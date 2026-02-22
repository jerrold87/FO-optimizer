import pandas as pd
import streamlit as st

from blend_fo import run_optimization


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
st.set_page_config(page_title="Fuel Oil Blending Optimizer", layout="wide")
st.title("Fuel Oil Blending Optimizer")
st.caption(
    "Upload your Fuel Oil input Excel file to solve the blending LP and download results. "
    "Tank min/max are in Metric Tons (MT). Viscosity uses Refutas Blending Number; "
    "Pour Point uses Pour Point Blending Index — both keep the LP fully linear."
)

uploaded = st.file_uploader("Upload input Excel (.xlsx)", type=["xlsx"])

if uploaded:
    if st.button("Run Optimization", type="primary"):
        with st.spinner("Solving…"):
            try:
                st.session_state["fo_opt_results"] = run_optimization(uploaded)
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
            st.dataframe(
                summary_display.style
                    .format("{:,.3f}", subset=["Mass (MT)", "Volume (m³)"])
                    .format("${:,.2f}", subset=["Value ($)", "Cost ($)", "Profit ($)"]),
                use_container_width=True,
            )

            # ----------------------------------------------------------
            # Components Usage
            # ----------------------------------------------------------
            st.subheader("Components Usage")
            st.dataframe(df.style
                    .format("{:,.3f}", subset=["Mass_MT", "Volume_m3"])
                    .format("${:,.2f}", subset=["Unit_Cost", "Total_Cost"]),
                use_container_width=True,
            )
            #st.download_button(
            #    label="Download Results as CSV",
            #    data=df.to_csv(index=False),
            #    file_name="fuel_oil_results.csv",
            #    mime="text/csv",
            #)

            # ----------------------------------------------------------
            # Blend Specifications
            # ----------------------------------------------------------
            st.subheader("Blend Specifications")
            grades = list(df_specs_display["Grade"].unique())
            tabs = st.tabs(grades)
            for tab, grade in zip(tabs, grades):
                with tab:
                    grade_df = (
                        df_specs_display[df_specs_display["Grade"] == grade]
                        .reset_index(drop=True)[["Property", "Blended", "Min", "Max"]]
                    )
                    st.dataframe(
                        grade_df.set_index("Property")[["Blended", "Min", "Max"]]
                            .style.apply(_highlight_spec, axis=1),
                        use_container_width=True,
                    )

            # ----------------------------------------------------------
            # Tank Summary
            # ----------------------------------------------------------
            st.subheader("Tank Summary")
            vol_cols = ["Before (Min)", "Before (Max)", "Used (Blend)", "After (Min)", "After (Max)"]
            st.dataframe(
                df_comp_summary.style.format("{:,.3f}", subset=vol_cols),
                use_container_width=True,
            )



        else:
            st.error(f"Solver returned status: **{status}**. No optimal solution found.")
            st.info(
                "Check that your input data is feasible — production demand, "
                "component inventory, and quality spec constraints must all be "
                "simultaneously satisfiable."
            )
