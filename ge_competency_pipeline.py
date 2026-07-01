from pathlib import Path
import pandas as pd

input_path = Path("engineering_competency_gap_dataset_v2.xlsx")

workbook = pd.read_excel(input_path, sheet_name=None)

for sheet_name, df in workbook.items():
    print(sheet_name, df.shape)
    
required_sheets = [
    "employees",
    "competency_model",
    "assessments",
    "training_actions",
    "evidence_log",
    "access_audit",
    "competency_targets",
]

missing_sheets = []

for sheet in required_sheets:
    if sheet not in workbook:
        missing_sheets.append(sheet)

if missing_sheets:
    raise ValueError(f"Missing required sheets: {missing_sheets}")

print("All required sheets are present.")

def clean_column_name(column_name):
    column_name = str(column_name)
    column_name = column_name.strip()
    column_name = column_name.lower()
    column_name = column_name.replace(" ", "_")
    column_name = column_name.replace("-", "_")
    column_name = column_name.replace("/", "_")
    return column_name

for sheet_name, df in workbook.items():
    df.columns = [clean_column_name(column) for column in df.columns]

#   
def clean_text_value(value):
    if pd.isna(value):
        return value

    value = str(value)
    value = value.strip()
    value = " ".join(value.split())
    return value

for sheet_name, df in workbook.items():
    text_columns = df.select_dtypes(include=["object", "string"]).columns

    for column in text_columns:
        df[column] = df[column].apply(clean_text_value)
        
print("Text values standardised.")
       
if "employees" in workbook:
    employees = workbook["employees"]

    if "account_status" in employees.columns:
        employees["account_status"] = (
            employees["account_status"]
            .str.lower()
            .map({
                "active": "Active",
                "inactive": "Inactive",
            })
            .fillna(employees["account_status"])
        )
        
def convert_date_columns(df, date_columns):
    for column in date_columns:
        if column in df.columns:
            df[column] = pd.to_datetime(df[column], errors="coerce")

    return df

date_columns_by_sheet = {
    "employees": ["start_date"],
    "assessments": ["assessment_date"],
    "training_actions": ["target_date"],
    "assessment_history": ["review_date"],
    "evidence_log": ["submitted_date"],
    "access_audit": ["last_login"],
}

for sheet_name, date_columns in date_columns_by_sheet.items():
    if sheet_name in workbook:
        workbook[sheet_name] = convert_date_columns(workbook[sheet_name], date_columns)

print("Date columns converted.")