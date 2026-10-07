from __future__ import annotations

from pathlib import Path
import pandas as pd

try:
    import xlsxwriter  # type: ignore[reportMissingImports]
except ImportError:  # pragma: no cover - handled at runtime when Excel output is requested.
    xlsxwriter = None


PROJECT_DIR = Path(__file__).resolve().parent
RAW_PATH = PROJECT_DIR / "healthcare_dataset.csv"
CLEANED_PATH = PROJECT_DIR / "cleaned_healthcare_data.csv"
OUTPUT_PATH = PROJECT_DIR / "Hospital_Patient_Analytics_Dashboard.xlsx"


def clean_healthcare_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    df = raw_df.copy()

    text_cols = [
        "Name",
        "Gender",
        "Blood Type",
        "Medical Condition",
        "Doctor",
        "Hospital",
        "Insurance Provider",
        "Medication",
        "Test Results",
        "Admission Type",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    if "Name" in df.columns:
        df["Name"] = df["Name"].str.title()
    if "Gender" in df.columns:
        df["Gender"] = df["Gender"].str.title()
    if "Blood Type" in df.columns:
        df["Blood Type"] = df["Blood Type"].str.upper()
    if "Medical Condition" in df.columns:
        df["Medical Condition"] = df["Medical Condition"].str.title()
    if "Doctor" in df.columns:
        df["Doctor"] = df["Doctor"].str.title()
    if "Hospital" in df.columns:
        df["Hospital"] = df["Hospital"].str.title()
    if "Insurance Provider" in df.columns:
        df["Insurance Provider"] = df["Insurance Provider"].str.title()
    if "Medication" in df.columns:
        df["Medication"] = df["Medication"].str.title()
    if "Test Results" in df.columns:
        df["Test Results"] = df["Test Results"].str.title()
    if "Admission Type" in df.columns:
        df["Admission Type"] = df["Admission Type"].str.title()

    df = df.drop_duplicates().copy()

    for col in ["Age", "Room Number", "Billing Amount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    date_cols = ["Date of Admission", "Discharge Date"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if "Date of Admission" in df.columns:
        df = df.dropna(subset=["Date of Admission"])
    if "Discharge Date" in df.columns:
        df = df.dropna(subset=["Discharge Date"])

    if "Age" in df.columns:
        df = df[df["Age"].between(0, 120)]
    if "Billing Amount" in df.columns:
        df = df[df["Billing Amount"] >= 0]
    if "Age" in df.columns:
        df["Age"] = df["Age"].astype(int)
    if "Room Number" in df.columns:
        df["Room Number"] = df["Room Number"].astype(int)

    if "Length of Stay" not in df.columns:
        df["Length of Stay"] = (df["Discharge Date"] -
                                df["Date of Admission"]).dt.days
    else:
        df["Length of Stay"] = pd.to_numeric(
            df["Length of Stay"], errors="coerce")

    if "Length of Stay" in df.columns:
        df = df[df["Length of Stay"] >= 0]
        df["Length of Stay"] = df["Length of Stay"].astype(int)

    df = df.reset_index(drop=True)
    return df


def write_data_sheet(ws, df: pd.DataFrame, start_row: int = 0, start_col: int = 0):
    headers = list(df.columns)
    for col_index, header in enumerate(headers):
        ws.write(start_row, start_col + col_index, header)
    for row_idx, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for col_idx, value in enumerate(row):
            ws.write(row_idx, start_col + col_idx, value)


def create_excel_dashboard():
    raw_df = pd.read_csv(RAW_PATH)
    cleaned_df = clean_healthcare_data(raw_df)
    cleaned_df.to_csv(CLEANED_PATH, index=False)

    workbook = xlsxwriter.Workbook(str(OUTPUT_PATH))
    workbook.set_properties({
        "title": "Hospital Patient Analytics Dashboard",
        "subject": "Healthcare analysis",
        "author": "AI Data Analytics Developer",
        "company": "Project",
        "category": "Analytics",
    })

    header_fmt = workbook.add_format({
        "bold": True,
        "font_color": "#FFFFFF",
        "bg_color": "#1F4E78",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "font_size": 10,
    })
    cell_fmt = workbook.add_format({"border": 1, "font_size": 10})
    num_fmt = workbook.add_format(
        {"border": 1, "num_format": "#,##0", "font_size": 10})
    currency_fmt = workbook.add_format(
        {"border": 1, "num_format": '$#,##0.00', "font_size": 10})
    percent_fmt = workbook.add_format(
        {"border": 1, "num_format": "0.0%", "font_size": 10})
    date_fmt = workbook.add_format(
        {"border": 1, "num_format": "dd-mmm-yyyy", "font_size": 10})
    title_fmt = workbook.add_format({
        "bold": True,
        "font_size": 18,
        "font_color": "#1F4E78",
        "align": "center",
    })
    kpi_label_fmt = workbook.add_format({
        "bold": True,
        "font_size": 9,
        "font_color": "#FFFFFF",
        "bg_color": "#2F75B5",
        "align": "center",
        "valign": "vcenter",
    })
    kpi_value_base = {
        "bold": True,
        "font_size": 16,
        "font_color": "#FFFFFF",
        "bg_color": "#3B7FBA",
        "align": "center",
        "valign": "vcenter",
    }
    kpi_count_fmt = workbook.add_format(
        {**kpi_value_base, "num_format": "#,##0"})
    kpi_age_fmt = workbook.add_format(
        {**kpi_value_base, "num_format": '0.0 "years"'})
    kpi_currency_fmt = workbook.add_format(
        {**kpi_value_base, "num_format": '"₹"#,##0'})
    kpi_stay_fmt = workbook.add_format(
        {**kpi_value_base, "num_format": '0.0 "days"'})

    raw_ws = workbook.add_worksheet("Raw_Data")
    write_data_sheet(raw_ws, raw_df, 0, 0)
    raw_ws.autofilter(0, 0, len(raw_df), len(raw_df.columns) - 1)
    raw_ws.freeze_panes(1, 0)

    clean_ws = workbook.add_worksheet("Cleaned_Data")
    write_data_sheet(clean_ws, cleaned_df, 0, 0)
    clean_ws.autofilter(0, 0, len(cleaned_df), len(cleaned_df.columns) - 1)
    clean_ws.freeze_panes(1, 0)

    calc_ws = workbook.add_worksheet("Calculations")
    calc_ws.set_column("A:B", 28)
    calc_ws.set_column("C:F", 18)
    calc_ws.merge_range("A1:B1", "KPI Summary", header_fmt)
    metrics = [
        ("Total Patients", '=COUNTA(\'Cleaned_Data\'!$A:$A)-1'),
        ("Average Age", '=AVERAGE(\'Cleaned_Data\'!$B:$B)'),
        ("Average Billing Amount", '=AVERAGE(\'Cleaned_Data\'!$J:$J)'),
        ("Average Length of Stay", '=AVERAGE(\'Cleaned_Data\'!$P:$P)'),
        ("Total Billing Amount", '=SUM(\'Cleaned_Data\'!$J:$J)'),
        ("Maximum Billing Amount", '=MAX(\'Cleaned_Data\'!$J:$J)'),
        ("Minimum Billing Amount", '=MIN(\'Cleaned_Data\'!$J:$J)'),
        ("Number of Hospitals", '=COUNTA(\'Cleaned_Data\'!$H:$H)-1'),
        ("Number of Doctors", '=COUNTA(\'Cleaned_Data\'!$G:$G)-1'),
        ("Number of Medical Conditions",
         '=SUMPRODUCT((\'Cleaned_Data\'!$E$2:$E$100000<>\"\")/COUNTIF(\'Cleaned_Data\'!$E$2:$E$100000,\'Cleaned_Data\'!$E$2:$E$100000&\"\")))'),
    ]

    calc_ws.write("A3", "Metric", header_fmt)
    calc_ws.write("B3", "Value", header_fmt)
    for idx, (label, formula) in enumerate(metrics, start=4):
        calc_ws.write(f"A{idx}", label, cell_fmt)
        calc_ws.write_formula(f"B{idx}", formula, cell_fmt)

    # Create histogram/bin summary tables for billing and age distributions
    bin_headers = ["Bucket", "Count"]
    billing_bins = [0, 10000, 20000, 30000, 40000, 50000, 60000]
    billing_labels = ["0-10K", "10K-20K",
                      "20K-30K", "30K-40K", "40K-50K", "50K-60K+"]
    age_bins = [0, 20, 30, 40, 50, 60, 70, 80, 90, 100]
    age_labels = ["0-19", "20-29", "30-39", "40-49",
                  "50-59", "60-69", "70-79", "80-89", "90+"]

    calc_ws.write("D3", "Billing Distribution", header_fmt)
    calc_ws.write("D4", "Bucket", header_fmt)
    calc_ws.write("E4", "Count", header_fmt)
    for i, (label, lower, upper) in enumerate([
        ("0-10K", 0, 10000),
        ("10K-20K", 10000, 20000),
        ("20K-30K", 20000, 30000),
        ("30K-40K", 30000, 40000),
        ("40K-50K", 40000, 50000),
        ("50K-60K+", 50000, 1000000),
    ], start=5):
        calc_ws.write(f"D{i}", label, cell_fmt)
        calc_ws.write_formula(
            f"E{i}", f'=COUNTIFS(\'Cleaned_Data\'!$J$2:$J$100000, ">={lower}", \'Cleaned_Data\'!$J$2:$J$100000, "<{upper}")', num_fmt)

    calc_ws.write("H3", "Age Distribution", header_fmt)
    calc_ws.write("H4", "Bucket", header_fmt)
    calc_ws.write("I4", "Count", header_fmt)
    for i, (label, lower, upper) in enumerate([
        ("0-19", 0, 20),
        ("20-29", 20, 30),
        ("30-39", 30, 40),
        ("40-49", 40, 50),
        ("50-59", 50, 60),
        ("60-69", 60, 70),
        ("70-79", 70, 80),
        ("80-89", 80, 90),
        ("90+", 90, 120),
    ], start=5):
        calc_ws.write(f"H{i}", label, cell_fmt)
        calc_ws.write_formula(
            f"I{i}", f'=COUNTIFS(\'Cleaned_Data\'!$B$2:$B$100000, ">={lower}", \'Cleaned_Data\'!$B$2:$B$100000, "<{upper}")', num_fmt)

    # Correlation matrix helper table
    calc_ws.write("L3", "Correlation Analysis", header_fmt)
    calc_ws.write("L4", "Variable", header_fmt)
    calc_ws.write("M4", "Age", header_fmt)
    calc_ws.write("N4", "Billing", header_fmt)
    calc_ws.write("O4", "Room", header_fmt)
    calc_ws.write("P4", "Length of Stay", header_fmt)
    row_labels = ["Age", "Billing", "Room", "Length of Stay"]
    data_cols = ["B", "J", "K", "P"]
    for r_idx, label in enumerate(row_labels, start=5):
        calc_ws.write(f"L{r_idx}", label, cell_fmt)
    for c_idx, col in enumerate(["B", "J", "K", "P"], start=13):
        for r_idx, row_label in enumerate(row_labels, start=5):
            target_row = 5 + row_labels.index(row_label)
            ref1 = f"'Cleaned_Data'!${row_labels[0] if False else 'B'}$2:$B$100000"

    # Fill the correlation matrix manually with formulas using actual Excel references
    corr_map = {
        ("Age", "Age"): '=CORREL(\'Cleaned_Data\'!$B$2:$B$100000,\'Cleaned_Data\'!$B$2:$B$100000)',
        ("Age", "Billing"): '=CORREL(\'Cleaned_Data\'!$B$2:$B$100000,\'Cleaned_Data\'!$J$2:$J$100000)',
        ("Age", "Room"): '=CORREL(\'Cleaned_Data\'!$B$2:$B$100000,\'Cleaned_Data\'!$K$2:$K$100000)',
        ("Age", "Length of Stay"): '=CORREL(\'Cleaned_Data\'!$B$2:$B$100000,\'Cleaned_Data\'!$P$2:$P$100000)',
        ("Billing", "Age"): '=CORREL(\'Cleaned_Data\'!$J$2:$J$100000,\'Cleaned_Data\'!$B$2:$B$100000)',
        ("Billing", "Billing"): '=CORREL(\'Cleaned_Data\'!$J$2:$J$100000,\'Cleaned_Data\'!$J$2:$J$100000)',
        ("Billing", "Room"): '=CORREL(\'Cleaned_Data\'!$J$2:$J$100000,\'Cleaned_Data\'!$K$2:$K$100000)',
        ("Billing", "Length of Stay"): '=CORREL(\'Cleaned_Data\'!$J$2:$J$100000,\'Cleaned_Data\'!$P$2:$P$100000)',
        ("Room", "Age"): '=CORREL(\'Cleaned_Data\'!$K$2:$K$100000,\'Cleaned_Data\'!$B$2:$B$100000)',
        ("Room", "Billing"): '=CORREL(\'Cleaned_Data\'!$K$2:$K$100000,\'Cleaned_Data\'!$J$2:$J$100000)',
        ("Room", "Room"): '=CORREL(\'Cleaned_Data\'!$K$2:$K$100000,\'Cleaned_Data\'!$K$2:$K$100000)',
        ("Room", "Length of Stay"): '=CORREL(\'Cleaned_Data\'!$K$2:$K$100000,\'Cleaned_Data\'!$P$2:$P$100000)',
        ("Length of Stay", "Age"): '=CORREL(\'Cleaned_Data\'!$P$2:$P$100000,\'Cleaned_Data\'!$B$2:$B$100000)',
        ("Length of Stay", "Billing"): '=CORREL(\'Cleaned_Data\'!$P$2:$P$100000,\'Cleaned_Data\'!$J$2:$J$100000)',
        ("Length of Stay", "Room"): '=CORREL(\'Cleaned_Data\'!$P$2:$P$100000,\'Cleaned_Data\'!$K$2:$K$100000)',
        ("Length of Stay", "Length of Stay"): '=CORREL(\'Cleaned_Data\'!$P$2:$P$100000,\'Cleaned_Data\'!$P$2:$P$100000)',
    }
    value_positions = {
        "Age": 13,
        "Billing": 14,
        "Room": 15,
        "Length of Stay": 16,
    }
    for row_label in row_labels:
        calc_ws.write(f"L{5 + row_labels.index(row_label)}",
                      row_label, cell_fmt)
        for col_label in row_labels:
            formula = corr_map[(row_label, col_label)]
            calc_ws.write_formula(
                f"{chr(77 + row_labels.index(col_label))}{5 + row_labels.index(row_label)}", formula, cell_fmt)

    # Pivot-like summary tables
    pivot_ws = workbook.add_worksheet("Pivot_Tables")
    pivot_ws.set_column("A:B", 24)
    pivot_ws.set_column("D:E", 24)
    pivot_ws.set_column("G:H", 24)
    pivot_ws.set_column("J:K", 24)
    pivot_ws.set_column("M:N", 24)
    pivot_ws.set_column("P:Q", 24)

    med_condition_list = sorted(
        cleaned_df["Medical Condition"].unique().tolist())
    gender_list = sorted(cleaned_df["Gender"].unique().tolist())
    admission_list = sorted(cleaned_df["Admission Type"].unique().tolist())
    insurance_list = sorted(cleaned_df["Insurance Provider"].unique().tolist())
    tests_list = sorted(cleaned_df["Test Results"].unique().tolist())
    first_month = cleaned_df["Date of Admission"].min(
    ).to_period("M").to_timestamp()
    last_month = cleaned_df["Date of Admission"].max(
    ).to_period("M").to_timestamp()
    month_labels = [
        month.strftime("%b-%Y")
        for month in pd.date_range(start=first_month, end=last_month, freq="MS")
    ]

    sections = [
        ("Patient count by Medical Condition", "A1",
         med_condition_list, "=COUNTIF('Cleaned_Data'!$E:$E,A2)"),
        ("Patient count by Gender", "D1", gender_list,
         "=COUNTIF('Cleaned_Data'!$C:$C,D2)"),
        ("Patient count by Admission Type", "G1",
         admission_list, "=COUNTIF('Cleaned_Data'!$L:$L,G2)"),
        ("Patient count by Insurance Provider", "J1",
         insurance_list, "=COUNTIF('Cleaned_Data'!$I:$I,J2)"),
        ("Patient count by Test Results", "M1",
         tests_list, "=COUNTIF('Cleaned_Data'!$O:$O,M2)"),
    ]
    for title, start_cell, values, formula_pattern in sections:
        col = start_cell[0]
        row = int(start_cell[1:])
        pivot_ws.merge_range(
            f"{col}{row}:{chr(ord(col)+1)}{row}", title, header_fmt)
        pivot_ws.write(f"{col}{row+1}", "Category", header_fmt)
        pivot_ws.write(f"{chr(ord(col)+1)}{row+1}", "Count", header_fmt)
        for idx, value in enumerate(values, start=row+2):
            pivot_ws.write(f"{col}{idx}", value, cell_fmt)
            pivot_ws.write_formula(f"{chr(ord(col)+1)}{idx}", formula_pattern.replace("A2", f"A{idx}").replace(
                "D2", f"D{idx}").replace("G2", f"G{idx}").replace("J2", f"J{idx}").replace("M2", f"M{idx}"), num_fmt)

    pivot_ws.write("A20", "Average Billing by Medical Condition", header_fmt)
    pivot_ws.write("A21", "Medical Condition", header_fmt)
    pivot_ws.write("B21", "Average Billing", header_fmt)
    for idx, condition in enumerate(med_condition_list, start=22):
        pivot_ws.write(f"A{idx}", condition, cell_fmt)
        pivot_ws.write_formula(
            f"B{idx}", f'=AVERAGEIF(\'Cleaned_Data\'!$E:$E,A{idx},\'Cleaned_Data\'!$J:$J)', currency_fmt)

    pivot_ws.write(
        "D20", "Average Length of Stay by Medical Condition", header_fmt)
    pivot_ws.write("D21", "Medical Condition", header_fmt)
    pivot_ws.write("E21", "Avg Length of Stay", header_fmt)
    for idx, condition in enumerate(med_condition_list, start=22):
        pivot_ws.write(f"D{idx}", condition, cell_fmt)
        pivot_ws.write_formula(
            f"E{idx}", f'=AVERAGEIF(\'Cleaned_Data\'!$E:$E,D{idx},\'Cleaned_Data\'!$P:$P)', cell_fmt)

    pivot_ws.write("G20", "Monthly Admissions", header_fmt)
    pivot_ws.write("G21", "Month", header_fmt)
    pivot_ws.write("H21", "Admissions", header_fmt)
    for idx, month in enumerate(month_labels, start=22):
        pivot_ws.write(f"G{idx}", month, cell_fmt)
        pivot_ws.write_formula(
            f"H{idx}", f'=SUMPRODUCT((TEXT(\'Cleaned_Data\'!$F$2:$F$100000,\"mmm-yyyy\")=G{idx}))', num_fmt)

    pivot_ws.write("J20", "Billing by Medical Condition", header_fmt)
    pivot_ws.write("J21", "Medical Condition", header_fmt)
    pivot_ws.write("K21", "Total Billing", header_fmt)
    for idx, condition in enumerate(med_condition_list, start=22):
        pivot_ws.write(f"J{idx}", condition, cell_fmt)
        pivot_ws.write_formula(
            f"K{idx}", f'=SUMIF(\'Cleaned_Data\'!$E:$E,J{idx},\'Cleaned_Data\'!$J:$J)', currency_fmt)

    # Dashboard setup
    dashboard_ws = workbook.add_worksheet("Dashboard")
    dashboard_ws.set_tab_color("#2F75B5")
    dashboard_ws.hide_gridlines(2)
    dashboard_ws.set_zoom(85)
    dashboard_ws.set_column("A:A", 2)
    dashboard_ws.set_column("B:Q", 12)
    dashboard_ws.set_row(0, 30)
    dashboard_ws.set_row(1, 18)
    dashboard_ws.set_row(2, 8)
    dashboard_ws.set_row(3, 8)
    dashboard_ws.set_row(4, 24)
    dashboard_ws.set_row(5, 25)
    dashboard_ws.set_row(6, 25)
    dashboard_ws.set_row(7, 12)
    dashboard_ws.merge_range(
        "B1:Q2", "🏥 HOSPITAL PATIENT ANALYTICS DASHBOARD", title_fmt
    )

    kpi_positions = [
        ("B", "E", "Total Patients", "=Calculations!B4", kpi_count_fmt),
        ("F", "I", "Avg Age", "=Calculations!B5", kpi_age_fmt),
        ("J", "M", "Avg Billing", "=Calculations!B6", kpi_currency_fmt),
        ("N", "Q", "Avg Stay", "=Calculations!B7", kpi_stay_fmt),
    ]
    for first_col, last_col, label, formula, value_fmt in kpi_positions:
        dashboard_ws.merge_range(
            f"{first_col}5:{last_col}5", label, kpi_label_fmt
        )
        dashboard_ws.merge_range(
            f"{first_col}6:{last_col}7", formula, value_fmt
        )

    # Monthly Patient Admissions (line chart)
    line_chart = workbook.add_chart({"type": "line"})
    line_chart.add_series({
        "name": "Monthly Admissions",
        "categories": f"='Pivot_Tables'!$G$22:$G${21 + len(month_labels)}",
        "values": f"='Pivot_Tables'!$H$22:$H${21 + len(month_labels)}",
        "line": {"color": "#2F75B5", "width": 2},
        "marker": {"type": "circle", "size": 4},
    })
    line_chart.set_title({"name": "Monthly Admissions"})
    line_chart.set_x_axis({"name": "Month"})
    line_chart.set_y_axis({"name": "Admissions"})
    line_chart.set_legend({"position": "bottom"})
    line_chart.set_chartarea({
        "fill": {"color": "#FFFFFF"},
        "border": {"color": "#D9E2F3"},
    })
    dashboard_ws.insert_chart(
        "B9", line_chart, {"x_scale": 0.9, "y_scale": 0.85})

    # Patients by Medical Condition (bar)
    medical_bar = workbook.add_chart({"type": "bar"})
    medical_bar.add_series({
        "name": "Patients",
        "categories": f"='Pivot_Tables'!$A$3:$A${len(med_condition_list) + 2}",
        "values": f"='Pivot_Tables'!$B$3:$B${len(med_condition_list) + 2}",
        "fill": {"color": "#1F4E78"},
        "border": {"color": "#1F4E78"},
    })
    medical_bar.set_title({"name": "Patients by Medical Condition"})
    medical_bar.set_x_axis({"name": "Patients"})
    medical_bar.set_y_axis({"name": "Medical Condition"})
    medical_bar.set_chartarea({
        "fill": {"color": "#FFFFFF"},
        "border": {"color": "#D9E2F3"},
    })
    dashboard_ws.insert_chart(
        "J9", medical_bar, {"x_scale": 0.9, "y_scale": 0.85})

    # Gender distribution pie chart
    gender_pie = workbook.add_chart({"type": "pie"})
    gender_pie.add_series({
        "name": "Gender",
        "categories": f"='Pivot_Tables'!$D$3:$D${len(gender_list) + 2}",
        "values": f"='Pivot_Tables'!$E$3:$E${len(gender_list) + 2}",
        "data_labels": {"percentage": True, "category": True, "leader_lines": True},
    })
    gender_pie.set_title({"name": "Gender Distribution"})
    gender_pie.set_legend({"position": "right"})
    gender_pie.set_chartarea({
        "fill": {"color": "#FFFFFF"},
        "border": {"color": "#D9E2F3"},
    })
    dashboard_ws.insert_chart(
        "B30", gender_pie, {"x_scale": 0.9, "y_scale": 0.85})

    # Billing distribution column chart
    billing_dist = workbook.add_chart({"type": "column"})
    billing_dist.add_series({
        "name": "Count",
        "categories": "='Calculations'!$D$5:$D$10",
        "values": "='Calculations'!$E$5:$E$10",
        "fill": {"color": "#4F81BD"},
        "border": {"color": "#4F81BD"},
        "gap": 0,
    })
    billing_dist.set_title({"name": "Billing Amount Distribution"})
    billing_dist.set_x_axis({"name": "Billing Amount"})
    billing_dist.set_y_axis({"name": "Patients"})
    billing_dist.set_chartarea({
        "fill": {"color": "#FFFFFF"},
        "border": {"color": "#D9E2F3"},
    })
    dashboard_ws.insert_chart("J30", billing_dist, {
                              "x_scale": 0.9, "y_scale": 0.85})

    # Scatter chart - Length of Stay vs Billing Amount
    scatter = workbook.add_chart({"type": "scatter", "subtype": "xy"})
    scatter.add_series({
        "name": "Stay vs Billing",
        "categories": f"='Cleaned_Data'!$P$2:$P${len(cleaned_df) + 1}",
        "values": f"='Cleaned_Data'!$J$2:$J${len(cleaned_df) + 1}",
        "marker": {"type": "circle", "size": 4, "fill": {"color": "#A5A5A5"}},
        "trendline": {"type": "linear", "line": {"color": "#C00000", "width": 2}},
    })
    scatter.set_title({"name": "Length of Stay vs Billing Amount"})
    scatter.set_x_axis({"name": "Length of Stay (days)"})
    scatter.set_y_axis({"name": "Billing Amount"})
    scatter.set_chartarea({
        "fill": {"color": "#FFFFFF"},
        "border": {"color": "#D9E2F3"},
    })
    dashboard_ws.insert_chart("B51", scatter, {"x_scale": 1.8, "y_scale": 0.9})

    workbook.close()


if __name__ == "__main__":
    create_excel_dashboard()
    print(f"Workbook created: {OUTPUT_PATH}")
    print(f"Cleaned CSV saved: {CLEANED_PATH}")
