"""
Fuel Oil Blending LP Optimization
==================================
Decision variables : mass blended per component per grade  [Metric Tons, MT]
Volume derived     : vol (m³) = mass (MT) / density (t/m³)

Non-linear properties handled by linearising through an index:

  Viscosity  – Refutas Blending Number
                 RBN(ν) = 14.534 × ln(ln(ν + 0.8)) + 10.975   [valid for ν > 0.2 cSt]
                 Blend RBN linearly by volume fraction, back-convert for display:
                 ν_blend = exp(exp((RBN_blend − 10.975) / 14.534)) − 0.8

  Pour Point – Pour Point Blending Index
                 PPBI(t) = ((t + 273.15) / 273.15)^12.5   [t in °C]
                 Blend PPBI linearly by volume fraction, back-convert for display:
                 t_blend = 273.15 × (PPBI_blend^(1/12.5) − 1)

  Flash Point – Hu-Burns Blending Index
                 FPBI(t) = exp(-0.06 × (t × 9/5 + 32))   [t in °C, internally in °F]
                 Blend FPBI linearly by volume fraction, back-convert for display:
                 t_blend = (ln(FPBI_blend) / (-0.06) − 32) × 5/9
                 NOTE: FPBI decreases as T increases → LP inequality directions are reversed

Both approaches reduce the non-linear blending constraint to a linear LP constraint:
  Σ INDEX_i · vol_i  ≥/≤  INDEX_limit · Σ vol_i
which is linear in the mass decision variables (since vol_i = mass_i / density_i).
"""

import math

import pandas as pd
from pulp import PULP_CBC_CMD, LpMaximize, LpProblem, LpStatus, LpVariable, lpSum, value

# ---------------------------------------------------------------------------
# Non-linear blending index helpers
# ---------------------------------------------------------------------------

_REFUTAS_THRESHOLD = 0.2  # minimum viscosity (cSt) for which RBN is valid


def _refutas(v: float) -> float:
    """Refutas Blending Number. v must be > 0.2 cSt so inner ln > 0."""
    return 14.534 * math.log(math.log(v + 0.8)) + 10.975


def _refutas_inv(x: float) -> float:
    """Back-convert Refutas Blending Number to kinematic viscosity (cSt)."""
    return math.exp(math.exp((x - 10.975) / 14.534)) - 0.8


def _ppbi(t: float) -> float:
    """Pour Point Blending Index. t in °C. Returns 1.0 at 0°C."""
    return ((t + 273.15) / 273.15) ** 12.5


def _ppbi_inv(x: float) -> float:
    """Back-convert PPBI to pour point (°C)."""
    return 273.15 * (x ** (1.0 / 12.5) - 1.0)


def _fpbi(t_c: float) -> float:
    """Hu-Burns Flash Point Blending Index. t_c in °C."""
    return math.exp(-0.06 * (t_c * 9.0 / 5.0 + 32.0))


def _fpbi_inv(x: float) -> float:
    """Back-convert Hu-Burns FPBI to flash point (°C)."""
    return (math.log(x) / (-0.06) - 32.0) * 5.0 / 9.0


def _gcv(density: float, water: float, ash: float, sulfur: float) -> float:
    """
    Gross Calorific Value (BTU/lb) per ISO/ASTM base method.
    density in t/m³ (= g/cm³); water, ash, sulfur in mass %.
    GCV = ((51.916 − 8.792×ρ²) × (1 − 0.01×(W+A+S)) + 9.42×(S/100)) × 1000 / 2.326
    """
    base = 51.916 - 8.792 * density ** 2
    corrected = base * (1.0 - 0.01 * (water + ash + sulfur)) + 9.42 * (sulfur / 100.0)
    return corrected * 1000.0 / 2.326


# ---------------------------------------------------------------------------
# Main optimisation function
# ---------------------------------------------------------------------------

def load_input_tables(uploaded_file):
    """Read the three input sections from the Excel file.

    Returns (df_comp, df_grades, df_specs) each indexed as the optimizer expects:
        df_comp   – indexed by 'Tank Name'
        df_grades – indexed by 'Grade Name'
        df_specs  – indexed by 'Property'
    """
    df_raw = pd.read_excel(uploaded_file, sheet_name='input', header=None)

    def _find_row(sentinel: str) -> int:
        mask = df_raw.iloc[:, 0] == sentinel
        hits = df_raw.index[mask].tolist()
        if not hits:
            raise ValueError(f"Cannot find '{sentinel}' in column A of the input sheet.")
        return int(hits[0])

    comp_hdr  = _find_row('Tank Name')
    grade_hdr = _find_row('Grade Name')
    spec_hdr  = _find_row('Property')

    # Components: read from header row; stop at first blank Tank Name
    df_comp_raw = pd.read_excel(uploaded_file, sheet_name='input', skiprows=comp_hdr)
    na_mask = df_comp_raw['Tank Name'].isna()
    blank_idx = int(na_mask.idxmax()) if na_mask.any() else len(df_comp_raw)
    df_comp = df_comp_raw.iloc[:blank_idx].set_index('Tank Name')

    # Grades: read from header row; stop at first blank Grade Name
    df_grades_raw = pd.read_excel(uploaded_file, sheet_name='input', skiprows=grade_hdr)
    na_mask = df_grades_raw['Grade Name'].isna()
    blank_idx = int(na_mask.idxmax()) if na_mask.any() else len(df_grades_raw)
    df_grades = df_grades_raw.iloc[:blank_idx].set_index('Grade Name')

    # Specs: read from header row to end of sheet
    df_specs = (
        pd.read_excel(uploaded_file, sheet_name='input', skiprows=spec_hdr)
        .dropna(subset=['Property'])
        .set_index('Property')
    )

    # Drop any empty/unnamed columns produced by blank cells in the Excel sheet
    for df in (df_comp, df_grades, df_specs):
        unnamed = [c for c in df.columns if str(c).startswith('Unnamed')]
        df.drop(columns=unnamed, inplace=True)

    return df_comp, df_grades, df_specs


def run_optimization(uploaded_file=None, *,
                     preloaded_comp=None, preloaded_grades=None, preloaded_specs=None) -> tuple:
    """
    Run the Fuel Oil blending LP optimisation.

    Args:
        uploaded_file: File-like object or path to the FO input Excel file.
                       Not required when preloaded_* dataframes are supplied.
        preloaded_comp:   df_comp indexed by 'Tank Name'   (optional)
        preloaded_grades: df_grades indexed by 'Grade Name' (optional)
        preloaded_specs:  df_specs indexed by 'Property'   (optional)

    Returns:
        6-tuple:
            df_results       – Grade/Component blend table (empty if infeasible)
            total_profit     – float (0.0 if infeasible)
            status           – solver status string, e.g. 'Optimal'
            df_blend_specs   – blended property values vs. spec limits per grade
            df_grade_summary – per-grade mass, volume, value, cost, profit
            df_comp_summary  – per-component before/after inventory + properties
    """

    # ------------------------------------------------------------------
    # 1. Data loading
    # ------------------------------------------------------------------
    if preloaded_comp is not None and preloaded_grades is not None and preloaded_specs is not None:
        df_comp, df_grades, df_specs = preloaded_comp, preloaded_grades, preloaded_specs
    else:
        df_comp, df_grades, df_specs = load_input_tables(uploaded_file)

    # ------------------------------------------------------------------
    # 2. Derived constants
    # ------------------------------------------------------------------
    components = df_comp.index.tolist()
    grades     = df_grades.index.tolist()
    properties = df_specs.index.tolist()

    # Volume coefficients: vol (m³) = mass (MT) / density (t/m³)
    vc = {c: 1.0 / df_comp.loc[c, 'Density'] for c in components}

    # Pre-compute Refutas Blending Number and Pour Point Blending Index for each component
    rfn_comp  = {c: _refutas(df_comp.loc[c, 'Viscosity']) for c in components}
    ppbi_comp = {c: _ppbi(df_comp.loc[c, 'Pour'])         for c in components}
    fpbi_comp = {c: _fpbi(df_comp.loc[c, 'Flash'])        for c in components}

    # Pre-compute per-component CCAI only when it appears in the spec sheet
    ccai_comp = (
        {
            c: df_comp.loc[c, 'Density'] * 1000
               - 141 * math.log10(math.log10(df_comp.loc[c, 'Viscosity'] + 0.85))
               - 81
            for c in components
        }
        if 'CCAI' in properties else {}
    )

    # Pre-compute per-component GCV only when it appears in the spec sheet
    gcv_comp = (
        {
            c: _gcv(df_comp.loc[c, 'Density'], df_comp.loc[c, 'Water'],
                    df_comp.loc[c, 'Ash'],     df_comp.loc[c, 'Sulfur'])
            for c in components
        }
        if 'GCV (btu/lb)' in properties else {}
    )

    # Identify non-linear properties (handled separately)
    NL_PROPS = {'Viscosity', 'Pour', 'Flash'}

    # ------------------------------------------------------------------
    # 3. Build LP
    # ------------------------------------------------------------------
    prob = LpProblem("FuelOil_Blending_Optimization", LpMaximize)

    # Decision variables: mass of component c blended into grade g  [MT]
    blend = LpVariable.dicts("m", (grades, components), lowBound=0)

    # Objective: Maximise Revenue − Cost  ($/MT basis)
    revenue = lpSum(
        df_grades.loc[g, 'Price'] * lpSum(blend[g][c] for c in components)
        for g in grades
    )
    costs = lpSum(
        df_comp.loc[c, 'Cost'] * lpSum(blend[g][c] for g in grades)
        for c in components
    )
    prob += revenue - costs

    for g in grades:
        col_prefix = g[:3]  # e.g. 'TAP' for 'TAPAR'

        # volumetric total for grade g (LP expression, linear in blend vars)
        total_vol_g = lpSum(vc[c] * blend[g][c] for c in components)

        # Constraint 1: Production demand [MT]
        total_mass_g = lpSum(blend[g][c] for c in components)
        prob += total_mass_g >= df_grades.loc[g, 'Min Production']
        prob += total_mass_g <= df_grades.loc[g, 'Max Production']

        # Constraint 3: Quality specs
        for p in properties:
            spec_min = df_specs.loc[p, f'{col_prefix} Min']
            spec_max = df_specs.loc[p, f'{col_prefix} Max']

            if p == 'Viscosity':
                # --- Refutas constraint (volume-weighted) ---
                rfn_expr = lpSum(rfn_comp[c] * vc[c] * blend[g][c] for c in components)
                if pd.notnull(spec_min) and spec_min > _REFUTAS_THRESHOLD:
                    prob += rfn_expr >= _refutas(spec_min) * total_vol_g
                if pd.notnull(spec_max):
                    prob += rfn_expr <= _refutas(spec_max) * total_vol_g

            elif p == 'Pour':
                # --- PPBI constraint (volume-weighted) ---
                ppbi_expr = lpSum(ppbi_comp[c] * vc[c] * blend[g][c] for c in components)
                if pd.notnull(spec_min):
                    prob += ppbi_expr >= _ppbi(spec_min) * total_vol_g
                if pd.notnull(spec_max):
                    prob += ppbi_expr <= _ppbi(spec_max) * total_vol_g

            elif p == 'Flash':
                # --- Hu-Burns constraint (volume-weighted, reversed inequalities) ---
                fpbi_expr = lpSum(fpbi_comp[c] * vc[c] * blend[g][c] for c in components)
                if pd.notnull(spec_min):
                    prob += fpbi_expr <= _fpbi(spec_min) * total_vol_g  # FPBI↓ as T↑
                if pd.notnull(spec_max):
                    prob += fpbi_expr >= _fpbi(spec_max) * total_vol_g  # FPBI↓ as T↑

            elif p == 'Density':
                # --- Density: volume-weighted ---
                actual = lpSum(df_comp.loc[c, p] * vc[c] * blend[g][c] for c in components)
                if pd.notnull(spec_min):
                    prob += actual >= spec_min * total_vol_g
                if pd.notnull(spec_max):
                    prob += actual <= spec_max * total_vol_g
            elif p == 'CCAI':
                # --- CCAI: volume-weighted blend of per-component CCAI values ---
                actual = lpSum(ccai_comp[c] * vc[c] * blend[g][c] for c in components)
                if pd.notnull(spec_min):
                    prob += actual >= spec_min * total_vol_g
                if pd.notnull(spec_max):
                    prob += actual <= spec_max * total_vol_g
            elif p == 'GCV (btu/lb)':
                # --- GCV: mass-weighted blend of per-component GCV values (BTU/lb) ---
                actual = lpSum(gcv_comp[c] * blend[g][c] for c in components)
                if pd.notnull(spec_min):
                    prob += actual >= spec_min * total_mass_g
                if pd.notnull(spec_max):
                    prob += actual <= spec_max * total_mass_g
            else:
                # --- Other linear properties: mass-weighted ---
                actual = lpSum(df_comp.loc[c, p] * blend[g][c] for c in components)
                if pd.notnull(spec_min):
                    prob += actual >= spec_min * total_mass_g
                if pd.notnull(spec_max):
                    prob += actual <= spec_max * total_mass_g

    # Constraint 2: Component availability [MT]
    for c in components:
        total_used = lpSum(blend[g][c] for g in grades)
        prob += total_used >= df_comp.loc[c, 'Min']
        prob += total_used <= df_comp.loc[c, 'Max']

    # ------------------------------------------------------------------
    # 4. Solve
    # ------------------------------------------------------------------
    prob.solve(PULP_CBC_CMD(msg=False))
    status = LpStatus[prob.status]

    # ------------------------------------------------------------------
    # 5. Collect results
    # ------------------------------------------------------------------
    if status == 'Optimal':

        # Blend results table
        results = []
        for g in grades:
            for c in components:
                mass = value(blend[g][c])
                if mass and mass > 1e-6:
                    vol = mass * vc[c]
                    results.append({
                        "Grade":          g,
                        "Tank":           c,
                        "Mass_MT":        round(mass, 3),
                        "Volume_m3":      round(vol,  3),
                        "Unit_Cost":      df_comp.loc[c, 'Cost'],
                        "Total_Cost":     round(mass * df_comp.loc[c, 'Cost'], 2),
                    })
        df_results = pd.DataFrame(results)
        total_profit = value(prob.objective)

        # Per-grade summary and blended specs (single pass)
        spec_rows    = []
        summary_rows = []

        for g in grades:
            col_prefix = g[:3]

            mass_vals = {c: value(blend[g][c]) for c in components}
            vol_vals  = {c: mass_vals[c] * vc[c] for c in components}

            total_mass = sum(mass_vals.values())
            total_vol  = sum(vol_vals.values())
            total_cost = sum(df_comp.loc[c, 'Cost'] * mass_vals[c] for c in components)
            total_val  = df_grades.loc[g, 'Price'] * total_mass

            summary_rows.append({
                "Grade":         g,
                "Total_Mass_MT": round(total_mass, 3),
                "Total_Vol_m3":  round(total_vol,  3),
                "Total_Profit":  round(total_val - total_cost, 2),
                "Total_Value":   round(total_val,  2),
                "Total_Cost":    round(total_cost, 2),
            })

            for p in properties:
                if total_vol > 0:
                    if p == 'Viscosity':
                        rfn_blend = sum(rfn_comp[c] * vol_vals[c] for c in components) / total_vol
                        blended_val = round(_refutas_inv(rfn_blend), 4)
                    elif p == 'Pour':
                        ppbi_blend = sum(ppbi_comp[c] * vol_vals[c] for c in components) / total_vol
                        blended_val = round(_ppbi_inv(ppbi_blend), 4)
                    elif p == 'Flash':
                        fpbi_blend = sum(fpbi_comp[c] * vol_vals[c] for c in components) / total_vol
                        blended_val = round(_fpbi_inv(fpbi_blend), 4)
                    elif p == 'Density':
                        blended_val = round(
                            sum(df_comp.loc[c, p] * vol_vals[c] for c in components) / total_vol, 4
                        )
                    elif p == 'CCAI':
                        blended_val = round(
                            sum(ccai_comp[c] * vol_vals[c] for c in components) / total_vol, 1
                        )
                    elif p == 'GCV (btu/lb)':
                        blended_val = round(
                            sum(gcv_comp[c] * mass_vals[c] for c in components) / total_mass, 1
                        )
                    else:
                        blended_val = round(
                            sum(df_comp.loc[c, p] * mass_vals[c] for c in components) / total_mass, 4
                        )
                else:
                    blended_val = 0.0

                spec_rows.append({
                    "Grade":    g,
                    "Property": p,
                    "Blended":  blended_val,
                    "Min":      df_specs.loc[p, f'{col_prefix} Min'],
                    "Max":      df_specs.loc[p, f'{col_prefix} Max'],
                })


        df_blend_specs   = pd.DataFrame(spec_rows)
        df_grade_summary = pd.DataFrame(summary_rows)

        # Component before/after inventory
        comp_rows = []
        for c in components:
            total_used = sum(value(blend[g][c]) for g in grades)
            orig_min = df_comp.loc[c, 'Min']
            orig_max = df_comp.loc[c, 'Max']
            row = {
                #"Cost ($/MT)":   df_comp.loc[c, 'Cost'],
                "Before (Min)":    orig_min,
                "Before (Max)":    orig_max,
                "Used (Blend)":     round(total_used, 3),
                "After (Min)":     round(max(0.0, orig_min - total_used), 3),
                "After (Max)":     round(orig_max - total_used, 3),
            }

            #for p in properties:
            #    row[p] = df_comp.loc[c, p]
            comp_rows.append(row)

        df_comp_summary = pd.DataFrame(comp_rows, index=components)
        df_comp_summary.index.name = "Tank"

    else:
        df_results       = pd.DataFrame()
        total_profit     = 0.0
        df_blend_specs   = pd.DataFrame()
        df_grade_summary = pd.DataFrame()
        df_comp_summary  = pd.DataFrame()

    return df_results, total_profit, status, df_blend_specs, df_grade_summary, df_comp_summary


# ---------------------------------------------------------------------------
# CLI entry-point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import os
    import sys

    file_path = "input-FO.xlsx"
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        sys.exit(1)

    df_results, profit, status, _, __, ___ = run_optimization(file_path)

    print("\n" + "=" * 40)
    print(f"STATUS: {status}")
    if status == "Optimal":
        print(f"TOTAL PROFIT: ${profit:,.2f}")
        df_results.to_csv("fuel_oil_results.csv", index=False)
        print("Results exported to 'fuel_oil_results.csv'")
        print(df_results.to_string(index=False))
    else:
        print("No optimal solution found.")
    print("=" * 40)
