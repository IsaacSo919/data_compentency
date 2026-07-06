from pathlib import Path
import pandas as pd

input_path = Path("engineering_competency_gap_dataset_v2.xlsx")

workbook = pd.read_excel(input_path, sheet_name=None)

for sheet_name, df in workbook.items(): #df means data_frame, i.e  a table(pandas' excel sheet)
    print(sheet_name, df.shape)# df.shape returns a tuple representing the dimensionality of the DataFrame. It gives you the number of rows and columns in the DataFrame. The first element of the tuple is the number of rows, and the second element is the number of columns.

required_sheets = [
    "employees",
    "competency_model",
    "assessments",
    "training_actions",
    "evidence_log",
    "access_audit",
    "competency_targets",
]

missing_sheets = [] # an empty array for missing sheet

for sheet in required_sheets: # this loop checks all required sheets.
    if sheet not in workbook:
        missing_sheets.append(sheet)

if missing_sheets: # this logic is if (missing_sheets is true(not empty)) then raise an error with the missing sheet names.
    raise ValueError(f"Missing required sheets: {missing_sheets}")

print("All required sheets are present.")

def clean_column_name(column_name):
    column_name = str(column_name)
    column_name = column_name.strip() #strip() method removes any leading and trailing whitespace characters from the string. e.g. if the column name is "  Employee ID  ", strip() will remove the spaces and return "Employee ID".
    column_name = column_name.lower() #lowercase
    column_name = column_name.replace(" ", "_") #snake_case: replaces spaces with underscores
    column_name = column_name.replace("-", "_") #snake_case: replaces hyphens with underscores
    column_name = column_name.replace("/", "_") #snake_case: replaces forward slashes with underscores
    return column_name

for sheet_name, df in workbook.items(): #For each sheet in the workbook, this loop iterates through the DataFrame and applies the clean_column_name function to each column name. The cleaned column names are then assigned back to the DataFrame's columns attribute.
    df.columns = [clean_column_name(column) for column in df.columns]

#   
def clean_text_value(value):
    if pd.isna(value): #isna() function checks if the value is NaN (Not a Number) or missing. If the value is NaN, it returns the value as is without any further processing.
        return value
    
    value = str(value) # convert the value to a string using str(value). 
    value = value.strip() #strip() method removes any leading and trailing whitespace characters from the string. e.g. if the value is "  Hello World  ", strip() will remove the spaces and return "Hello World".
    value = " ".join(value.split())# The split() method splits the string into a list of words based on whitespace, and then the join() method combines the words back into a single string with a single space between each word. This effectively removes any extra spaces between words.
    return value

for sheet_name, df in workbook.items():
    text_columns = df.select_dtypes(include=["object", "string"]).columns # Select all columns with object or string data types

    for column in text_columns:
        df[column] = df[column].apply(clean_text_value)
        
print("Text values standardised.")
       
if "employees" in workbook:
    employees = workbook["employees"]

    if "account_status" in employees.columns: #account_status is a column in employees sheet. This logic checks if the account_status column exists in the employees DataFrame. If it does, the code proceeds to standardize the values in that column.
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

for index, row in assessments.iterrows():
    employee_id = row["employee_id"]

    if employee_id not in valid_employee_ids: # data quality report
        add_issue(
            source_table="assessments",
            record_id=row["assessment_id"],
            issue_type="invalid_employee_id",
            severity="High",
            description=f"Assessment references employee_id '{employee_id}', but this employee does not exist in the employees table.",
            recommended_action="Review the assessment employee_id or add the missing employee record.",
        )

competency_model = workbook["competency_model"]

valid_competency_ids = set(competency_model["competency_id"])

for index, row in assessments.iterrows():
    competency_id = row["competency_id"]

    if competency_id not in valid_competency_ids: # data quality report
        add_issue(
            source_table="assessments",
            record_id=row["assessment_id"],
            issue_type="invalid_competency_id",
            severity="High",
            description=f"Assessment references competency_id '{competency_id}', but this competency does not exist in the competency_model table.",
            recommended_action="Review the assessment competency_id or add the missing competency model record.",
        )
        
# 1. evidence_log assessment_id check
evidence_log = workbook["evidence_log"]

valid_assessment_ids = set(assessments["assessment_id"])

for index, row in evidence_log.iterrows():
    assessment_id = row["assessment_id"]

    if assessment_id not in valid_assessment_ids:
        add_issue(
            source_table="evidence_log",
            record_id=row["evidence_id"],
            issue_type="invalid_evidence_assessment_id",
            severity="High",
            description=f"Evidence references assessment_id '{assessment_id}', but this assessment does not exist in the assessments table.",
            recommended_action="Review the evidence assessment_id or link it to a valid assessment record.",
        )
        
# 2. training_actions employee_id check        
training_actions = workbook["training_actions"]

valid_employee_ids = set(employees["employee_id"])

for index, row in training_actions.iterrows():
    employee_id = row["employee_id"]

    if employee_id not in valid_employee_ids:
        add_issue(
            source_table="training_actions",
            record_id=row["action_id"],
            issue_type="invalid_training_employee_id",
            severity="High",
            description=f"Training action references employee_id '{employee_id}', but this employee does not exist in the employees table.",
            recommended_action="Review the training action employee_id or add the missing employee record.",
        )
        
#3. training_actions competency_id check
'''
自然語言：
    每個 training action 應該對應一個 valid competency。
'''
valid_competency_ids = set(competency_model["competency_id"])

for index, row in training_actions.iterrows():
    competency_id = row["competency_id"]

    if competency_id not in valid_competency_ids:
        add_issue(
            source_table="training_actions",
            record_id=row["action_id"],
            issue_type="invalid_training_competency_id",
            severity="High",
            description=f"Training action references competency_id '{competency_id}', but this competency does not exist in the competency_model table.",
            recommended_action="Review the training action competency_id or add the missing competency model record.",
        )

print(f"Relationship checks completed. Total issues found: {len(data_quality_issues)}")

# 4. missing evidence check
evidence_log = workbook["evidence_log"]

assessment_ids_with_evidence = set(evidence_log["assessment_id"])

for index, row in assessments.iterrows():
    assessment_id = row["assessment_id"]
    evidence_status = row["evidence_status"]

    if evidence_status in ["Submitted", "Verified"]:
        if assessment_id not in assessment_ids_with_evidence:
            add_issue(
                source_table="assessments",
                record_id=assessment_id,
                issue_type="missing_evidence",
                severity="Medium",
                description=f"Assessment '{assessment_id}' has evidence_status '{evidence_status}', but no matching evidence record exists in evidence_log.",
                recommended_action="Check whether evidence was not uploaded, incorrectly linked, or missing from the evidence log.",
            )
print(assessments["evidence_status"].value_counts(dropna=False))

# 5. stale assessment check
for index, row in assessments.iterrows():
    assessment_age_days = row["assessment_age_days"]

    if pd.notna(assessment_age_days) and assessment_age_days > 365:
        add_issue(
            source_table="assessments",
            record_id=row["assessment_id"],
            issue_type="stale_assessment",
            severity="Medium",
            description=f"Assessment is {assessment_age_days} days old and may need review.",
            recommended_action="Review and refresh the assessment if competency evidence is no longer current.",
        )
print(f"Stale assessment checks completed. Total issues found: {len(data_quality_issues)}")

# 6. inactive employee record check
access_audit = workbook["access_audit"]

inactive_employee_ids = set(
    employees.loc[
        employees["account_status"] == "Inactive",
        "employee_id"
    ]
)

for index, row in access_audit.iterrows():
    employee_id = row["employee_id"]

    if employee_id in inactive_employee_ids:
        add_issue(
            source_table="access_audit",
            record_id=row["audit_id"],
            issue_type="inactive_employee_access_record",
            severity="High",
            description=f"Inactive employee '{employee_id}' still appears in the access audit records.",
            recommended_action="Review whether access should be removed, disabled, or confirmed as closed.",
        )
print(f"Inactive employee access checks completed. Total issues found: {len(data_quality_issues)}")