
import pandas as pd
from pulp import PULP_CBC_CMD, LpMaximize, LpProblem, LpStatus, LpVariable, lpSum, value


_KPA_PER_PSI = 6.89476  # 1 psi = 6.89476 kPa


def _rvpbi(p_psi: float) -> float:
    """RVP Blending Index (Power 1.25 method). p_psi must be in psi."""
    return p_psi ** 1.25


def _rvpbi_inv(x: float) -> float:
    """Back-convert RVP Blending Index to RVP (psi)."""
    return x ** 0.8


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
    Run the gasoline blending LP optimization.

    Args:
        uploaded_file: A file-like object or path to the input Excel file.
                       Not required when preloaded_* dataframes are supplied.
        preloaded_comp:   df_comp indexed by 'Tank Name'  (optional)
        preloaded_grades: df_grades indexed by 'Grade Name'    (optional)
        preloaded_specs:  df_specs indexed by 'Property'       (optional)

    Returns:
        6-tuple:
            df_results       – Grade/Component blend table (empty if infeasible)
            total_profit     – float (0.0 if infeasible)
            status           – solver status string, e.g. 'Optimal'
            df_blend_specs   – blended property values vs. spec limits per grade
            df_grade_summary – per-grade volume, value, cost, profit
            df_comp_summary  – per-component before/after inventory
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

    # Mass conversion coefficient: (bbl → m³) × (Density kg/L − 0.0011) → MT per bbl
    mass_coeff = {c: 0.158987295 * (df_comp.loc[c, 'Density'] - 0.0011) for c in components}

    # Pre-compute per-component AKI (avg of RON & MON) only when it appears in the spec sheet
    aki_comp = (
        {c: (df_comp.loc[c, 'RON'] + df_comp.loc[c, 'MON']) / 2.0 for c in components}
        if 'AKI' in properties else {}
    )

    # Pre-compute per-component Sensitivity (RON − MON) only when it appears in the spec sheet
    sensitivity_comp = (
        {c: df_comp.loc[c, 'RON'] - df_comp.loc[c, 'MON'] for c in components}
        if 'Sensitivity' in properties else {}
    )

    # Pre-compute per-component DriveIndex (T10×1.5 + T50×3 + T90) only when it appears in the spec sheet
    driveindex_comp = (
        {c: df_comp.loc[c, 'T10'] * 1.5 + df_comp.loc[c, 'T50'] * 3.0 + df_comp.loc[c, 'T90']
         for c in components}
        if 'DriveIndex' in properties else {}
    )

    # Pre-compute per-component RVP Blending Index — PSI and kPa variants (formula always in PSI)
    rvpbi_psi_comp = (
        {c: _rvpbi(df_comp.loc[c, 'RVP_PSI']) for c in components}
        if 'RVP_PSI' in properties else {}
    )
    rvpbi_kpa_comp = (
        {c: _rvpbi(df_comp.loc[c, 'RVP_kPa'] / _KPA_PER_PSI) for c in components}
        if 'RVP_kPa' in properties else {}
    )

    # ------------------------------------------------------------------
    # 3. Build LP Model
    # ------------------------------------------------------------------
    prob = LpProblem("Gasoline_Blending_Optimization", LpMaximize)

    # Decision Variables: barrels of component 'c' used in grade 'g'
    blend = LpVariable.dicts("v", (grades, components), lowBound=0)

    # Objective: Maximize Revenue - Cost
    revenue = lpSum(df_grades.loc[g, 'Price'] * lpSum(blend[g][c] for c in components) for g in grades)
    costs   = lpSum(df_comp.loc[c, 'Cost']   * lpSum(blend[g][c] for g in grades)     for c in components)
    prob += revenue - costs

    # Constraint 1: Production Demand (min/max barrels per grade)
    for g in grades:
        total_vol = lpSum(blend[g][c] for c in components)
        prob += total_vol >= df_grades.loc[g, 'Min Production']
        prob += total_vol <= df_grades.loc[g, 'Max Production']

    # Constraint 2: Component Availability (inventory limits)
    for c in components:
        total_used = lpSum(blend[g][c] for g in grades)
        prob += total_used >= df_comp.loc[c, 'Min']
        prob += total_used <= df_comp.loc[c, 'Max']

    # Constraint 3: Quality Specifications (all linear, volume-weighted)
    for g in grades:
        total_vol  = lpSum(blend[g][c] for c in components)
        col_prefix = g[:3]
        for p in properties:
            min_limit  = df_specs.loc[p, f'{col_prefix} Min']
            max_limit  = df_specs.loc[p, f'{col_prefix} Max']
            if p == 'AKI':
                actual_val = lpSum(aki_comp[c] * blend[g][c] for c in components)
                if pd.notnull(min_limit):
                    prob += actual_val >= min_limit * total_vol
                if pd.notnull(max_limit):
                    prob += actual_val <= max_limit * total_vol
            elif p == 'Sensitivity':
                actual_val = lpSum(sensitivity_comp[c] * blend[g][c] for c in components)
                if pd.notnull(min_limit):
                    prob += actual_val >= min_limit * total_vol
                if pd.notnull(max_limit):
                    prob += actual_val <= max_limit * total_vol
            elif p == 'DriveIndex':
                actual_val = lpSum(driveindex_comp[c] * blend[g][c] for c in components)
                if pd.notnull(min_limit):
                    prob += actual_val >= min_limit * total_vol
                if pd.notnull(max_limit):
                    prob += actual_val <= max_limit * total_vol
            elif p == 'RVP_PSI':
                rvpbi_expr = lpSum(rvpbi_psi_comp[c] * blend[g][c] for c in components)
                if pd.notnull(min_limit):
                    prob += rvpbi_expr >= _rvpbi(min_limit) * total_vol
                if pd.notnull(max_limit):
                    prob += rvpbi_expr <= _rvpbi(max_limit) * total_vol
            elif p == 'RVP_kPa':
                rvpbi_expr = lpSum(rvpbi_kpa_comp[c] * blend[g][c] for c in components)
                if pd.notnull(min_limit):
                    prob += rvpbi_expr >= _rvpbi(min_limit / _KPA_PER_PSI) * total_vol
                if pd.notnull(max_limit):
                    prob += rvpbi_expr <= _rvpbi(max_limit / _KPA_PER_PSI) * total_vol
            elif p in ('O2', 'Sulfur'):
                # Mass-weighted: weight by MT = mass_coeff[c] * bbl
                total_mass = lpSum(mass_coeff[c] * blend[g][c] for c in components)
                actual_val = lpSum(df_comp.loc[c, p] * mass_coeff[c] * blend[g][c] for c in components)
                if pd.notnull(min_limit):
                    prob += actual_val >= min_limit * total_mass
                if pd.notnull(max_limit):
                    prob += actual_val <= max_limit * total_mass
            else:
                actual_val = lpSum(df_comp.loc[c, p] * blend[g][c] for c in components)
                if pd.notnull(min_limit):
                    prob += actual_val >= min_limit * total_vol
                if pd.notnull(max_limit):
                    prob += actual_val <= max_limit * total_vol

    # ------------------------------------------------------------------
    # 4. Solve
    # ------------------------------------------------------------------
    prob.solve(PULP_CBC_CMD(msg=False))
    status = LpStatus[prob.status]

    # ------------------------------------------------------------------
    # 5. Collect results
    # ------------------------------------------------------------------
    if status == 'Optimal':
        results = []
        for g in grades:
            for c in components:
                vol = value(blend[g][c])
                if vol and vol > 1e-6:
                    results.append({
                        "Grade":      g,
                        "Tank":       c,
                        "Volume_bbl": round(vol, 2),
                        "Mass_MT":    round(vol * mass_coeff[c], 3),
                        "Unit_Cost":  df_comp.loc[c, 'Cost'],
                        "Total_Cost": round(vol * df_comp.loc[c, 'Cost'], 2),
                    })
        df_results   = pd.DataFrame(results)
        total_profit = value(prob.objective)

        # Per-grade summary and blended specs (single pass)
        spec_rows    = []
        summary_rows = []

        for g in grades:
            col_prefix = g[:3]
            vol_vals   = {c: value(blend[g][c]) for c in components}
            total_vol  = sum(vol_vals.values())
            total_cost = sum(df_comp.loc[c, 'Cost'] * vol_vals[c] for c in components)
            total_val  = df_grades.loc[g, 'Price'] * total_vol

            total_mass = sum(vol_vals[c] * mass_coeff[c] for c in components)
            summary_rows.append({
                "Grade":            g,
                "Total_Volume_bbl": round(total_vol,  2),
                "Total_Mass_MT":    round(total_mass, 3),
                "Total_Value":      round(total_val,  2),
                "Total_Cost":       round(total_cost, 2),
                "Total_Profit":     round(total_val - total_cost, 2),
            })

            for p in properties:
                if total_vol > 0:
                    if p == 'AKI':
                        blended_val = sum(aki_comp[c] * vol_vals[c] for c in components) / total_vol
                    elif p == 'Sensitivity':
                        blended_val = sum(sensitivity_comp[c] * vol_vals[c] for c in components) / total_vol
                    elif p == 'DriveIndex':
                        blended_val = sum(driveindex_comp[c] * vol_vals[c] for c in components) / total_vol
                    elif p == 'RVP_PSI':
                        rvpbi_blend = sum(rvpbi_psi_comp[c] * vol_vals[c] for c in components) / total_vol
                        blended_val = _rvpbi_inv(rvpbi_blend)
                    elif p == 'RVP_kPa':
                        rvpbi_blend = sum(rvpbi_kpa_comp[c] * vol_vals[c] for c in components) / total_vol
                        blended_val = _rvpbi_inv(rvpbi_blend) * _KPA_PER_PSI
                    elif p in ('O2', 'Sulfur'):
                        mass_vals  = {c: vol_vals[c] * mass_coeff[c] for c in components}
                        total_mass = sum(mass_vals.values())
                        blended_val = (
                            sum(df_comp.loc[c, p] * mass_vals[c] for c in components) / total_mass
                            if total_mass > 0 else 0.0
                        )
                    else:
                        blended_val = sum(df_comp.loc[c, p] * vol_vals[c] for c in components) / total_vol
                else:
                    blended_val = 0.0
                spec_rows.append({
                    "Grade":    g,
                    "Property": p,
                    "Blended":  round(blended_val, 4),
                    "Min":      df_specs.loc[p, f'{col_prefix} Min'],
                    "Max":      df_specs.loc[p, f'{col_prefix} Max'],
                })

        df_blend_specs   = pd.DataFrame(spec_rows)
        df_grade_summary = pd.DataFrame(summary_rows)

        # Component before/after inventory
        comp_rows = []
        for c in components:
            total_used = sum(value(blend[g][c]) for g in grades)
            orig_min   = df_comp.loc[c, 'Min']
            orig_max   = df_comp.loc[c, 'Max']
            comp_rows.append({
                "Before (Min)": orig_min,
                "Before (Max)": orig_max,
                "Used (Blend)": round(total_used, 2),
                "After (Min)":  round(max(0.0, orig_min - total_used), 2),
                "After (Max)":  round(orig_max - total_used, 2),
            })

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

    file_path = r"C:\Users\jerro\OneDrive\Desktop\Python\Datafile\input-gasoline.v2.xlsx"
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        sys.exit(1)

    df_results, profit, status, _, __, ___ = run_optimization(file_path)

    print("\n" + "=" * 30)
    print(f"STATUS: {status}")
    if status == "Optimal":
        print(f"TOTAL PROFIT: ${profit:,.2f}")
        df_results.to_csv("gasoline_results.csv", index=False)
        print("Results exported to 'gasoline_results.csv'")
        print(df_results)
    else:
        print("No optimal solution found.")
    print("=" * 30)
