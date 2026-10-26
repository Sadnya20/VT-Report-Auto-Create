
import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Employee Attendance Report",
    page_icon="📊",
    layout="wide"
)

# Required report columns
REQUIRED_COLUMNS = [
    "Date",
    "Employee Name",
    "Designation",
    "Working_With",
    "Covered_Beat",
    "Attendance",
    "Market Objective",
    "Attendance Time",
    "Atten. Remarks"
]

# Only these designations are allowed
ALLOWED_DESIGNATIONS = [
    "ASM", "ASE", "SDM", "RSM", "TSE", "TSA", "NSM"
]

st.title("📊 Employee Attendance Report")
st.write("Upload an Excel or CSV file to generate your report.")

uploaded_file = st.file_uploader(
    "Upload Excel or CSV",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:

    try:
        filename = uploaded_file.name.lower()

        # Read uploaded file
        if filename.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            excel_file = pd.ExcelFile(uploaded_file)

            if len(excel_file.sheet_names) > 1:
                sheet = st.selectbox(
                    "Select worksheet",
                    excel_file.sheet_names
                )
            else:
                sheet = excel_file.sheet_names[0]

            df = pd.read_excel(
                uploaded_file,
                sheet_name=sheet
            )

        # Clean accidental spaces in column headers
        df.columns = df.columns.astype(str).str.strip()

        st.success("File uploaded successfully!")

        # Validate required columns
        missing_columns = [
            col for col in REQUIRED_COLUMNS
            if col not in df.columns
        ]

        if missing_columns:
            st.error(
                "Required columns missing: "
                + ", ".join(missing_columns)
            )
            st.stop()

        # Keep only required columns
        report = df[REQUIRED_COLUMNS].copy()

        # Standardize designation before filtering
        report["Designation"] = (
            report["Designation"]
            .astype("string")
            .str.strip()
            .str.upper()
        )

        # Keep only allowed designations
        report = report[
            report["Designation"].isin(
                ALLOWED_DESIGNATIONS
            )
        ].copy()

        # Optional interactive filters
        st.subheader("Filters")

        c1, c2 = st.columns(2)

        with c1:
            designations = st.multiselect(
                "Select Designation",
                ALLOWED_DESIGNATIONS,
                default=ALLOWED_DESIGNATIONS
            )

        with c2:
            employees = sorted(
                report["Employee Name"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            selected_employees = st.multiselect(
                "Select Employee",
                employees
            )

        # Apply selected designation filter
        report = report[
            report["Designation"].isin(designations)
        ]

        # Apply employee filter only when selected
        if selected_employees:
            report = report[
                report["Employee Name"]
                .astype(str)
                .isin(selected_employees)
            ]

        # Summary metrics
        st.subheader("Report Summary")

        c1, c2, c3 = st.columns(3)

        c1.metric("Source Rows", len(df))
        c2.metric("Filtered Rows", len(report))
        c3.metric(
            "Employees",
            report["Employee Name"].nunique()
        )

        st.subheader("Report Preview")

        st.dataframe(
            report,
            use_container_width=True,
            hide_index=True
        )

        # Create designation summary
        summary = (
            report.groupby("Designation", dropna=False)
            .size()
            .reset_index(name="Record Count")
        )

        # Generate Excel workbook
        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            report.to_excel(
                writer,
                sheet_name="Attendance Report",
                index=False
            )

            summary.to_excel(
                writer,
                sheet_name="Designation Summary",
                index=False
            )

            # Professional Excel formatting
            for ws in writer.book.worksheets:
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                ws.sheet_view.showGridLines = False

                for cell in ws[1]:
                    cell.font = cell.font.copy(
                        bold=True,
                        color="FFFFFF"
                    )
                    from openpyxl.styles import PatternFill
                    cell.fill = PatternFill(
                        fill_type="solid",
                        fgColor="17365D"
                    )

                for column_cells in ws.columns:
                    max_length = max(
                        len(str(cell.value or ""))
                        for cell in list(column_cells)[:500]
                    )
                    ws.column_dimensions[
                        column_cells[0].column_letter
                    ].width = min(max(max_length + 2, 12), 32)

        output.seek(0)

        st.download_button(
            label="📥 Download Formatted Excel Report",
            data=output.getvalue(),
            file_name="Employee_Attendance_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
            type="primary"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")
