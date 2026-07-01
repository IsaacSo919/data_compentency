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

data_quality_issues = []

def add_issue(source_table, record_id, issue_type, severity, description, recommended_action):
    data_quality_issues.append({
        "source_table": source_table,
        "record_id": record_id,
        "issue_type": issue_type,
        "severity": severity,
        "description": description,
        "recommended_action": recommended_action,
    })
    
assessments = workbook["assessments"]

for index, row in assessments.iterrows():
    current_level = row["current_level"]

    if pd.isna(current_level) or current_level < 1 or current_level > 5:
        add_issue(
            source_table="assessments",
            record_id=row["assessment_id"],
            issue_type="invalid_level_value",
            severity="High",
            description=f"Assessment has invalid current level: {current_level}",
            recommended_action="Review assessment level and correct it to a valid 1-5 value.",
        )
        
for index, row in assessments.iterrows():
    required_level = row["required_level"]

    if pd.isna(required_level) or required_level < 1 or required_level > 5:
        add_issue(
            source_table="assessments",
            record_id=row["assessment_id"],
            issue_type="invalid_required_level",
            severity="High",
            description=f"Assessment has invalid required level: {required_level}",
            recommended_action="Review required competency level and correct it to a valid 1-5 value.",
        )
        
print(f"Business rule checks completed. Issues found: {len(data_quality_issues)}")

print("Current level min:", assessments["current_level"].min())
print("Current level max:", assessments["current_level"].max())
print("Required level min:", assessments["required_level"].min())
print("Required level max:", assessments["required_level"].max())
print("Missing current level:", assessments["current_level"].isna().sum())
print("Missing required level:", assessments["required_level"].isna().sum())

#Below is the relationship check section:
employees = workbook["employees"]
assessments = workbook["assessments"]

valid_employee_ids = set(employees["employee_id"])
assessment_employee_ids = set(assessments["employee_id"])

invalid_assessment_employee_ids = assessment_employee_ids - valid_employee_ids

for employee_id in invalid_assessment_employee_ids:
    add_issue(
        source_table="assessments",
        record_id=employee_id,
        issue_type="broken_relationship",
        severity="High",
        description=f"Assessment references employee_id '{employee_id}', but this employee does not exist in the employees table.",
        recommended_action="Review the assessment employee_id or add the missing employee record.",
    )

print(f"Employee relationship checks completed. Invalid employee IDs found: {len(invalid_assessment_employee_ids)}")
competency_model = workbook["competency_model"]

valid_competency_ids = set(competency_model["competency_id"])
assessment_competency_ids = set(assessments["competency_id"])

invalid_assessment_competency_ids = assessment_competency_ids - valid_competency_ids

for competency_id in invalid_assessment_competency_ids:
    add_issue(
        source_table="assessments",
        record_id=competency_id,
        issue_type="broken_relationship",
        severity="High",
        description=f"Assessment references competency_id '{competency_id}', but this competency does not exist in the competency_model table.",
        recommended_action="Review the assessment competency_id or add the missing competency model record.",
    )

print(f"Competency relationship checks completed. Invalid competency IDs found: {len(invalid_assessment_competency_ids)}")