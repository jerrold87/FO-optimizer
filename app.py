import io

import pandas as pd
import streamlit as st

import blend_fo
import blend_gasoline


def _highlight_spec(row):
    """Colour the row red if the blended value violates the spec limit."""
    val, lo, hi = row["Blended"], row["Min"], row["Max"]
    color = ""
    if (pd.notnull(lo) and val < lo) or (pd.notnull(hi) and val > hi):
        color = "background-color: #f28b82"
    return [color] * len(row)


# ---------------------------------------------------------------------------
# Page layout
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Multi-Product Blending Optimizer", layout="wide")
st.title("Multi-Product Blending Optimizer")

product = st.radio("Select product to blend:", ["Fuel Oil", "Gasoline"], horizontal=True)

st.divider()

# ===========================================================================
# FUEL OIL
# ===========================================================================
if product == "Fuel Oil":
    st.caption(
        "Upload your Fuel Oil input Excel file to solve the blending LP and download results. "
        "Tank min/max are in Metric Tons (MT). Viscosity uses Refutas Blending Number; "
        "Pour Point uses Pour Point Blending Index; "
        "Flash Point uses Hu-Burns Blending Index — keeping the LP fully linear."
    )

    uploaded = st.file_uploader("Upload input Excel (.xlsx)", type=["xlsx"], key="fo_uploader")

    if uploaded:
        if st.session_state.get("fo_input_file") != uploaded.name:
            df_comp_raw, df_grades_raw, df_specs_raw = blend_fo.load_input_tables(uploaded)

            df_comp_disp = df_comp_raw.reset_index()
            df_comp_disp.insert(0, "Include", True)

            df_grades_disp = df_grades_raw.reset_index()
            df_grades_disp.insert(0, "Include", True)

            st.session_state["fo_input_raw"] = (
                df_comp_disp,
                df_grades_disp,
                df_specs_raw.reset_index(),
            )
            st.session_state["fo_input_file"] = uploaded.name

        df_comp_init, df_grades_init, df_specs_init = st.session_state["fo_input_raw"]

        with st.expander("Input Data (review & edit before running)", expanded=True):
            tab_comp, tab_grades, tab_specs = st.tabs(["Components", "Grades", "Specs"])

            with tab_comp:
                edited_comp = st.data_editor(
                    df_comp_init, num_rows="dynamic", use_container_width=True,
                    key=f"fo_editor_comp_{st.session_state.get('fo_editor_comp_ver', 0)}",
                    column_config={"Include": st.column_config.CheckboxColumn("Include", default=True)},
                )

            with tab_grades:
                edited_grades = st.data_editor(
                    df_grades_init, num_rows="dynamic", use_container_width=True,
                    key="fo_editor_grades",
                    column_config={"Include": st.column_config.CheckboxColumn("Include", default=True)},
                )

            with tab_specs:
                edited_specs = st.data_editor(
                    df_specs_init, num_rows="dynamic", use_container_width=True,
                    key="fo_editor_specs",
                )

        if st.button("Run Optimization", type="primary", key="fo_run"):
            with st.spinner("Solving…"):
                try:
                    df_comp_in = (
                        edited_comp[edited_comp["Include"]]
                        .drop(columns="Include")
                        .set_index("Tank Name")
                        .dropna(how="all")
                    )
                    df_grades_in = (
                        edited_grades[edited_grades["Include"]]
                        .drop(columns="Include")
                        .set_index("Grade Name")
                        .dropna(how="all")
                    )
                    df_specs_in = edited_specs.set_index("Property").dropna(how="all")
                    st.session_state["fo_opt_results"] = blend_fo.run_optimization(
                        preloaded_comp=df_comp_in,
                        preloaded_grades=df_grades_in,
                        preloaded_specs=df_specs_in,
                    )
                except Exception as e:
                    st.error(f"Failed to run model: {e}")
                    st.stop()

        if "fo_opt_results" in st.session_state:
            df, profit, status, df_specs_display, df_grade_summary, df_comp_summary = (
                st.session_state["fo_opt_results"]
            )

            if status == "Optimal":
                col1, col2 = st.columns(2)
                col1.metric("Solver Status", status)
                col2.metric("Total Profit", f"${profit:,.2f}")

                # ----------------------------------------------------------
                # Blend Summary
                # ----------------------------------------------------------
                st.subheader("Blend Summary")
                summary_display = df_grade_summary.set_index("Grade").copy()
                summary_display.columns = [
                    "Mass (MT)", "Volume (m³)", "Profit ($)", "Value ($)", "Cost ($)"
                ]
                summary_display["$ / Mass (MT)"] = summary_display["Cost ($)"] / summary_display["Mass (MT)"]
                summary_display = summary_display[["$ / Mass (MT)", "Mass (MT)", "Volume (m³)", "Profit ($)", "Value ($)", "Cost ($)"]]
                st.dataframe(
                    summary_display.style
                        .format("{:,.3f}", subset=["Mass (MT)", "Volume (m³)"])
                        .format("${:,.2f}", subset=["Value ($)", "Cost ($)", "Profit ($)"])
                        .format("${:,.2f}", subset=["$ / Mass (MT)"]),
                    width='stretch',
                )

                # ----------------------------------------------------------
                # Components Usage
                # ----------------------------------------------------------
                st.subheader("Components Usage")
                comp_display = df.copy()
                comp_display["% (MT)"] = comp_display["Mass_MT"] / comp_display.groupby("Grade")["Mass_MT"].transform("sum")
                cols = ["Grade", "Tank", "% (MT)", "Mass_MT", "Volume_m3", "Unit_Cost", "Total_Cost"]
                comp_display = comp_display[cols].rename(columns={
                    "Mass_MT":    "Mass (MT)",
                    "Volume_m3":  "Volume (m³)",
                    "Unit_Cost":  "Unit Cost ($)",
                    "Total_Cost": "Total Cost ($)",
                })
                comp_grades = list(comp_display["Grade"].unique())
                comp_tabs = st.tabs(comp_grades)
                for tab, grade in zip(comp_tabs, comp_grades):
                    with tab:
                        grade_comp = comp_display[comp_display["Grade"] == grade].drop(columns="Grade").reset_index(drop=True)
                        st.dataframe(
                            grade_comp.style
                                .format("{:,.3f}", subset=["Mass (MT)", "Volume (m³)"])
                                .format("${:,.2f}", subset=["Unit Cost ($)", "Total Cost ($)"])
                                .format("{:.1%}", subset=["% (MT)"]),
                            width='stretch',
                        )

                # ----------------------------------------------------------
                # Blend Specifications
                # ----------------------------------------------------------
                st.subheader("Blend Specifications")
                grades_list = list(df_specs_display["Grade"].unique())
                spec_tabs = st.tabs(grades_list)
                for tab, grade in zip(spec_tabs, grades_list):
                    with tab:
                        grade_df = (
                            df_specs_display[df_specs_display["Grade"] == grade]
                            .reset_index(drop=True)[["Property", "Blended", "Min", "Max"]]
                        )
                        st.dataframe(
                            grade_df.set_index("Property")[["Blended", "Min", "Max"]]
                                .style.apply(_highlight_spec, axis=1),
                            width='stretch',
                        )

                # ----------------------------------------------------------
                # Tank Summary
                # ----------------------------------------------------------
                st.subheader("Tank Summary")
                vol_cols = ["Before (Min)", "Before (Max)", "Used (Blend)", "After (Min)", "After (Max)"]
                st.dataframe(
                    df_comp_summary.style.format("{:,.3f}", subset=vol_cols),
                    width='stretch',
                )
                if st.button("↺ Roll balance: update Components Min/Max from Tank Summary", key="fo_roll"):
                    updated_comp = edited_comp.copy()
                    for tank, row in df_comp_summary.iterrows():
                        mask = updated_comp["Tank Name"] == tank
                        if mask.any():
                            updated_comp.loc[mask, "Min"] = row["After (Min)"]
                            updated_comp.loc[mask, "Max"] = row["After (Max)"]
                    _, grades_part, specs_part = st.session_state["fo_input_raw"]
                    st.session_state["fo_input_raw"] = (updated_comp, grades_part, specs_part)
                    st.session_state["fo_editor_comp_ver"] = st.session_state.get("fo_editor_comp_ver", 0) + 1
                    st.rerun()

                # ----------------------------------------------------------
                # Download
                # ----------------------------------------------------------
                st.divider()
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    summary_display.to_excel(writer, sheet_name="Blend Summary")
                    comp_display.to_excel(writer, sheet_name="Components Usage", index=False)
                    (
                        df_specs_display[["Grade", "Property", "Blended", "Min", "Max"]]
                        .to_excel(writer, sheet_name="Blend Specifications", index=False)
                    )
                    df_comp_summary.to_excel(writer, sheet_name="Tank Summary")

                scenario_name = st.text_input(
                    "Blend Plan",
                    placeholder="Enter a name for this blend plan…",
                    key="fo_scenario",
                )
                file_name = f"{scenario_name.strip()}.xlsx" if scenario_name.strip() else "fo_blend_plan.xlsx"
                st.download_button(
                    label="Download Results as Excel",
                    data=buf.getvalue(),
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="fo_download",
                )

            else:
                st.error(f"Solver returned status: **{status}**. No optimal solution found.")
                st.info(
                    "Check that your input data is feasible — production demand, "
                    "component inventory, and quality spec constraints must all be "
                    "simultaneously satisfiable."
                )

# ===========================================================================
# GASOLINE
# ===========================================================================
else:
    st.caption(
        "Upload your Gasoline input Excel file to solve the blending LP and download results. "
        "Tank min/max are in Barrels (bbl). RVP uses Blending Index Formula while all other quality properties blend linearly by volume fraction."
    )

    uploaded = st.file_uploader("Upload input Excel (.xlsx)", type=["xlsx"], key="gas_uploader")

    if uploaded:
        if st.session_state.get("gas_input_file") != uploaded.name:
            df_comp_raw, df_grades_raw, df_specs_raw = blend_gasoline.load_input_tables(uploaded)

            df_comp_disp = df_comp_raw.reset_index()
            df_comp_disp.insert(0, "Include", True)

            df_grades_disp = df_grades_raw.reset_index()
            df_grades_disp.insert(0, "Include", True)

            st.session_state["gas_input_raw"] = (
                df_comp_disp,
                df_grades_disp,
                df_specs_raw.reset_index(),
            )
            st.session_state["gas_input_file"] = uploaded.name

        df_comp_init, df_grades_init, df_specs_init = st.session_state["gas_input_raw"]

        with st.expander("Input Data (review & edit before running)", expanded=True):
            tab_comp, tab_grades, tab_specs = st.tabs(["Components", "Grades", "Specs"])

            with tab_comp:
                edited_comp = st.data_editor(
                    df_comp_init, num_rows="dynamic", use_container_width=True,
                    key=f"gas_editor_comp_{st.session_state.get('gas_editor_comp_ver', 0)}",
                    column_config={"Include": st.column_config.CheckboxColumn("Include", default=True)},
                )

            with tab_grades:
                edited_grades = st.data_editor(
                    df_grades_init, num_rows="dynamic", use_container_width=True,
                    key="gas_editor_grades",
                    column_config={"Include": st.column_config.CheckboxColumn("Include", default=True)},
                )

            with tab_specs:
                edited_specs = st.data_editor(
                    df_specs_init, num_rows="dynamic", use_container_width=True,
                    key="gas_editor_specs",
                )

        if st.button("Run Optimization", type="primary", key="gas_run"):
            with st.spinner("Solving…"):
                try:
                    df_comp_in = (
                        edited_comp[edited_comp["Include"]]
                        .drop(columns="Include")
                        .set_index("Tank Name")
                        .dropna(how="all")
                    )
                    df_grades_in = (
                        edited_grades[edited_grades["Include"]]
                        .drop(columns="Include")
                        .set_index("Grade Name")
                        .dropna(how="all")
                    )
                    df_specs_in = edited_specs.set_index("Property").dropna(how="all")
                    st.session_state["gas_opt_results"] = blend_gasoline.run_optimization(
                        preloaded_comp=df_comp_in,
                        preloaded_grades=df_grades_in,
                        preloaded_specs=df_specs_in,
                    )
                except Exception as e:
                    st.error(f"Failed to run model: {e}")
                    st.stop()

        if "gas_opt_results" in st.session_state:
            df, profit, status, df_specs_display, df_grade_summary, df_comp_summary = (
                st.session_state["gas_opt_results"]
            )

            if status == "Optimal":
                col1, col2 = st.columns(2)
                col1.metric("Solver Status", status)
                col2.metric("Total Profit", f"${profit:,.2f}")

                # ----------------------------------------------------------
                # Blend Summary
                # ----------------------------------------------------------
                st.subheader("Blend Summary")
                summary_display = df_grade_summary.set_index("Grade").copy()
                summary_display.columns = ["Volume (bbl)", "Value ($)", "Cost ($)", "Profit ($)"]
                summary_display["$ / Volume (bbl)"] = summary_display["Cost ($)"] / summary_display["Volume (bbl)"]
                summary_display = summary_display[["$ / Volume (bbl)", "Volume (bbl)", "Profit ($)", "Value ($)", "Cost ($)"]]
                st.dataframe(
                    summary_display.style
                        .format("{:,.2f}", subset=["Volume (bbl)"])
                        .format("${:,.2f}", subset=["Value ($)", "Cost ($)", "Profit ($)", "$ / Volume (bbl)"]),
                    width='stretch',
                )

                # ----------------------------------------------------------
                # Components Usage
                # ----------------------------------------------------------
                st.subheader("Components Usage")
                comp_display = df.copy()
                comp_display["% (bbl)"] = comp_display["Volume_bbl"] / comp_display.groupby("Grade")["Volume_bbl"].transform("sum")
                cols = ["Grade", "Tank", "% (bbl)", "Volume_bbl", "Unit_Cost", "Total_Cost"]
                comp_display = comp_display[cols].rename(columns={
                    "Volume_bbl": "Volume (bbl)",
                    "Unit_Cost":  "Unit Cost ($)",
                    "Total_Cost": "Total Cost ($)",
                })
                comp_grades = list(comp_display["Grade"].unique())
                comp_tabs = st.tabs(comp_grades)
                for tab, grade in zip(comp_tabs, comp_grades):
                    with tab:
                        grade_comp = comp_display[comp_display["Grade"] == grade].drop(columns="Grade").reset_index(drop=True)
                        st.dataframe(
                            grade_comp.style
                                .format("{:,.2f}", subset=["Volume (bbl)"])
                                .format("${:,.2f}", subset=["Unit Cost ($)", "Total Cost ($)"])
                                .format("{:.1%}", subset=["% (bbl)"]),
                            width='stretch',
                        )

                # ----------------------------------------------------------
                # Blend Specifications
                # ----------------------------------------------------------
                st.subheader("Blend Specifications")
                grades_list = list(df_specs_display["Grade"].unique())
                spec_tabs = st.tabs(grades_list)
                for tab, grade in zip(spec_tabs, grades_list):
                    with tab:
                        grade_df = (
                            df_specs_display[df_specs_display["Grade"] == grade]
                            .reset_index(drop=True)[["Property", "Blended", "Min", "Max"]]
                        )
                        st.dataframe(
                            grade_df.set_index("Property")[["Blended", "Min", "Max"]]
                                .style.apply(_highlight_spec, axis=1),
                            width='stretch',
                        )

                # ----------------------------------------------------------
                # Component Summary
                # ----------------------------------------------------------
                st.subheader("Component Summary")
                vol_cols = ["Before (Min)", "Before (Max)", "Used (Blend)", "After (Min)", "After (Max)"]
                st.dataframe(
                    df_comp_summary.style.format("{:,.2f}", subset=vol_cols),
                    width='stretch',
                )
                if st.button("↺ Roll balance: update Components Min/Max from Component Summary", key="gas_roll"):
                    updated_comp = edited_comp.copy()
                    for comp_name, row in df_comp_summary.iterrows():
                        mask = updated_comp["Tank Name"] == comp_name
                        if mask.any():
                            updated_comp.loc[mask, "Min"] = row["After (Min)"]
                            updated_comp.loc[mask, "Max"] = row["After (Max)"]
                    _, grades_part, specs_part = st.session_state["gas_input_raw"]
                    st.session_state["gas_input_raw"] = (updated_comp, grades_part, specs_part)
                    st.session_state["gas_editor_comp_ver"] = st.session_state.get("gas_editor_comp_ver", 0) + 1
                    st.rerun()

                # ----------------------------------------------------------
                # Download
                # ----------------------------------------------------------
                st.divider()
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    summary_display.to_excel(writer, sheet_name="Blend Summary")
                    comp_display.to_excel(writer, sheet_name="Components Usage", index=False)
                    (
                        df_specs_display[["Grade", "Property", "Blended", "Min", "Max"]]
                        .to_excel(writer, sheet_name="Blend Specifications", index=False)
                    )
                    df_comp_summary.to_excel(writer, sheet_name="Component Summary")

                scenario_name = st.text_input(
                    "Blend Plan",
                    placeholder="Enter a name for this blend plan…",
                    key="gas_scenario",
                )
                file_name = f"{scenario_name.strip()}.xlsx" if scenario_name.strip() else "gasoline_blend_plan.xlsx"
                st.download_button(
                    label="Download Results as Excel",
                    data=buf.getvalue(),
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="gas_download",
                )

            else:
                st.error(f"Solver returned status: **{status}**. No optimal solution found.")
                st.info(
                    "Check that your input data is feasible — production demand, "
                    "component inventory, and quality spec constraints must all be "
                    "simultaneously satisfiable."
                )
