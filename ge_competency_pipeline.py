from pathlib import Path

import pandas as pd


INPUT_PATH = Path("engineering_competency_gap_dataset_v2.xlsx")
OUTPUT_FOLDER = Path("outputs")

REQUIRED_SHEETS = [
    "employees",
    "competency_model",
    "assessments",
    "training_actions",
    "evidence_log",
    "access_audit",
    "competency_targets",
]

DATE_COLUMNS_BY_SHEET = {
    "employees": ["start_date"],
    "assessments": ["assessment_date"],
    "training_actions": ["target_date"],
    "assessment_history": ["review_date"],
    "evidence_log": ["submitted_date"],
    "access_audit": ["last_login"],
}


def clean_column_name(column_name):
    column_name = str(column_name)
    column_name = column_name.strip()
    column_name = column_name.lower()
    column_name = column_name.replace(" ", "_")
    column_name = column_name.replace("-", "_")
    column_name = column_name.replace("/", "_")
    return column_name


def clean_text_value(value):
    if pd.isna(value):
        return value

    value = str(value)
    value = value.strip()
    value = " ".join(value.split())
    return value


def load_workbook(input_path):
    workbook = pd.read_excel(input_path, sheet_name=None)

    for sheet_name, df in workbook.items():
        print(sheet_name, df.shape)

    return workbook


def validate_required_sheets(workbook, required_sheets):
    missing_sheets = []

    for sheet in required_sheets:
        if sheet not in workbook:
            missing_sheets.append(sheet)

    if missing_sheets:
        raise ValueError(f"Missing required sheets: {missing_sheets}")

    print("All required sheets are present.")


def standardise_column_names(workbook):
    for sheet_name, df in workbook.items():
        df.columns = [clean_column_name(column) for column in df.columns]

    print("Column names standardised.")


def standardise_text_values(workbook):
    for sheet_name, df in workbook.items():
        text_columns = df.select_dtypes(include=["object", "string"]).columns

        for column in text_columns:
            df[column] = df[column].apply(clean_text_value)

    print("Text values standardised.")


def standardise_status_values(workbook):
    employees = workbook["employees"]

    if "account_status" in employees.columns:
        employees["account_status"] = (
            employees["account_status"]
            .str.lower()
            .map(
                {
                    "active": "Active",
                    "inactive": "Inactive",
                }
            )
            .fillna(employees["account_status"])
        )

    print("Status values standardised.")


def convert_date_columns(workbook, date_columns_by_sheet):
    for sheet_name, date_columns in date_columns_by_sheet.items():
        if sheet_name in workbook:
            df = workbook[sheet_name]

            for column in date_columns:
                if column in df.columns:
                    df[column] = pd.to_datetime(df[column], errors="coerce")

    print("Date columns converted.")


def add_issue(data_quality_issues, source_table, record_id, issue_type, severity, description, recommended_action):
    data_quality_issues.append(
        {
            "source_table": source_table,
            "record_id": record_id,
            "issue_type": issue_type,
            "severity": severity,
            "description": description,
            "recommended_action": recommended_action,
        }
    )


def check_level_values(workbook, data_quality_issues):
    assessments = workbook["assessments"]

    for index, row in assessments.iterrows():
        current_level = row["current_level"]

        if pd.isna(current_level) or current_level < 1 or current_level > 5:
            add_issue(
                data_quality_issues,
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
                data_quality_issues,
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


def check_relationships(workbook, data_quality_issues):
    employees = workbook["employees"]
    competency_model = workbook["competency_model"]
    assessments = workbook["assessments"]
    evidence_log = workbook["evidence_log"]
    training_actions = workbook["training_actions"]

    valid_employee_ids = set(employees["employee_id"])
    valid_competency_ids = set(competency_model["competency_id"])
    valid_assessment_ids = set(assessments["assessment_id"])

    for index, row in assessments.iterrows():
        if row["employee_id"] not in valid_employee_ids:
            add_issue(
                data_quality_issues,
                source_table="assessments",
                record_id=row["assessment_id"],
                issue_type="invalid_employee_id",
                severity="High",
                description=f"Assessment references employee_id '{row['employee_id']}', but this employee does not exist in the employees table.",
                recommended_action="Review the assessment employee_id or add the missing employee record.",
            )

        if row["competency_id"] not in valid_competency_ids:
            add_issue(
                data_quality_issues,
                source_table="assessments",
                record_id=row["assessment_id"],
                issue_type="invalid_competency_id",
                severity="High",
                description=f"Assessment references competency_id '{row['competency_id']}', but this competency does not exist in the competency_model table.",
                recommended_action="Review the assessment competency_id or add the missing competency model record.",
            )

    for index, row in evidence_log.iterrows():
        if row["assessment_id"] not in valid_assessment_ids:
            add_issue(
                data_quality_issues,
                source_table="evidence_log",
                record_id=row["evidence_id"],
                issue_type="invalid_evidence_assessment_id",
                severity="High",
                description=f"Evidence references assessment_id '{row['assessment_id']}', but this assessment does not exist in the assessments table.",
                recommended_action="Review the evidence assessment_id or link it to a valid assessment record.",
            )

    for index, row in training_actions.iterrows():
        if row["employee_id"] not in valid_employee_ids:
            add_issue(
                data_quality_issues,
                source_table="training_actions",
                record_id=row["action_id"],
                issue_type="invalid_training_employee_id",
                severity="High",
                description=f"Training action references employee_id '{row['employee_id']}', but this employee does not exist in the employees table.",
                recommended_action="Review the training action employee_id or add the missing employee record.",
            )

        if row["competency_id"] not in valid_competency_ids:
            add_issue(
                data_quality_issues,
                source_table="training_actions",
                record_id=row["action_id"],
                issue_type="invalid_training_competency_id",
                severity="High",
                description=f"Training action references competency_id '{row['competency_id']}', but this competency does not exist in the competency_model table.",
                recommended_action="Review the training action competency_id or add the missing competency model record.",
            )

    print(f"Relationship checks completed. Total issues found: {len(data_quality_issues)}")


def check_missing_evidence(workbook, data_quality_issues):
    assessments = workbook["assessments"]
    evidence_log = workbook["evidence_log"]
    assessment_ids_with_evidence = set(evidence_log["assessment_id"])

    for index, row in assessments.iterrows():
        assessment_id = row["assessment_id"]
        evidence_status = row["evidence_status"]

        if evidence_status in ["Submitted", "Verified", "Partial"]:
            if assessment_id not in assessment_ids_with_evidence:
                add_issue(
                    data_quality_issues,
                    source_table="assessments",
                    record_id=assessment_id,
                    issue_type="missing_evidence",
                    severity="Medium",
                    description=f"Assessment '{assessment_id}' has evidence_status '{evidence_status}', but no matching evidence record exists in evidence_log.",
                    recommended_action="Check whether evidence was not uploaded, incorrectly linked, or missing from the evidence log.",
                )

    print(assessments["evidence_status"].value_counts(dropna=False))
    print(f"Missing evidence checks completed. Total issues found: {len(data_quality_issues)}")


def check_stale_assessments(workbook, data_quality_issues):
    assessments = workbook["assessments"]

    for index, row in assessments.iterrows():
        assessment_age_days = row["assessment_age_days"]

        if pd.notna(assessment_age_days) and assessment_age_days > 365:
            add_issue(
                data_quality_issues,
                source_table="assessments",
                record_id=row["assessment_id"],
                issue_type="stale_assessment",
                severity="Medium",
                description=f"Assessment is {assessment_age_days} days old and may need review.",
                recommended_action="Review and refresh the assessment if competency evidence is no longer current.",
            )

    print(f"Stale assessment checks completed. Total issues found: {len(data_quality_issues)}")


def check_inactive_employee_access(workbook, data_quality_issues):
    employees = workbook["employees"]
    access_audit = workbook["access_audit"]

    inactive_employee_ids = set(
        employees.loc[
            employees["account_status"] == "Inactive",
            "employee_id",
        ]
    )

    for index, row in access_audit.iterrows():
        employee_id = row["employee_id"]

        if employee_id in inactive_employee_ids:
            add_issue(
                data_quality_issues,
                source_table="access_audit",
                record_id=row["audit_id"],
                issue_type="inactive_employee_access_record",
                severity="High",
                description=f"Inactive employee '{employee_id}' still appears in the access audit records.",
                recommended_action="Review whether access should be removed, disabled, or confirmed as closed.",
            )

    print(f"Inactive employee access checks completed. Total issues found: {len(data_quality_issues)}")


def check_high_gap_without_training_action(workbook, data_quality_issues):
    assessments = workbook["assessments"]
    training_actions = workbook["training_actions"]

    training_action_pairs = set(
        zip(
            training_actions["employee_id"],
            training_actions["competency_id"],
        )
    )

    for index, row in assessments.iterrows():
        employee_id = row["employee_id"]
        competency_id = row["competency_id"]
        gap = row["gap"]

        if pd.notna(gap) and gap >= 2:
            if (employee_id, competency_id) not in training_action_pairs:
                add_issue(
                    data_quality_issues,
                    source_table="assessments",
                    record_id=row["assessment_id"],
                    issue_type="high_gap_without_training_action",
                    severity="High",
                    description=f"Employee '{employee_id}' has a competency gap of {gap} for competency '{competency_id}', but no matching training action exists.",
                    recommended_action="Create or review a training action to address this competency gap.",
                )

    print(f"High gap training action checks completed. Total issues found: {len(data_quality_issues)}")


def export_data_quality_issues(data_quality_issues, output_folder):
    issues_df = pd.DataFrame(data_quality_issues)

    output_folder.mkdir(exist_ok=True)
    issues_df.to_csv(output_folder / "data_quality_issues.csv", index=False)
    issues_df.to_excel(output_folder / "data_quality_issues.xlsx", index=False)

    print("Data quality issue files exported.")

    if not issues_df.empty:
        print(issues_df["issue_type"].value_counts())
    else:
        print("No data quality issues found.")

def build_summary_report(workbook, data_quality_issues):
    issues_df = pd.DataFrame(data_quality_issues)

    if issues_df.empty:
        print("No issues to summarise.")
        return {}

    summary = {}

    summary["issues_by_type"] = issues_df["issue_type"].value_counts()
    summary["issues_by_severity"] = issues_df["severity"].value_counts()
    summary["issues_by_source_table"] = issues_df["source_table"].value_counts()

    print("\nIssues by type:")
    print(summary["issues_by_type"])

    print("\nIssues by severity:")
    print(summary["issues_by_severity"])

    print("\nIssues by source table:")
    print(summary["issues_by_source_table"])

    return summary


def export_enriched_issue_report(workbook, data_quality_issues, output_folder):
    issues_df = pd.DataFrame(data_quality_issues)

    if issues_df.empty:
        print("No issues to enrich.")
        return

    assessments = workbook["assessments"]
    access_audit = workbook["access_audit"]

    assessment_context_columns = [
        "assessment_id",
        "employee_id",
        "employee_name",
        "team",
        "role_family",
        "competency_id",
        "competency_area",
        "competency_name",
        "gap",
        "criticality",
        "current_level",
        "required_level",
        "evidence_status",
        "assessment_age_days",
    ]

    access_context_columns = [
        "audit_id",
        "employee_id",
        "employee_name",
        "team",
        "permission_role",
        "tool_access",
        "last_login",
        "access_issue",
        "ticket_status",
        "ticket_age_days",
    ]

    assessment_issues = issues_df[issues_df["source_table"] == "assessments"].copy()
    access_issues = issues_df[issues_df["source_table"] == "access_audit"].copy()

    enriched_assessment_issues = assessment_issues.merge(
        assessments[assessment_context_columns],
        how="left",
        left_on="record_id",
        right_on="assessment_id",
    )

    enriched_access_issues = access_issues.merge(
        access_audit[access_context_columns],
        how="left",
        left_on="record_id",
        right_on="audit_id",
    )

    enriched_issues = pd.concat(
        [enriched_assessment_issues, enriched_access_issues],
        ignore_index=True,
        sort=False,
    )

    output_folder.mkdir(exist_ok=True)

    enriched_issues.to_csv(
        output_folder / "enriched_data_quality_issues.csv",
        index=False,
    )

    enriched_issues.to_excel(
        output_folder / "enriched_data_quality_issues.xlsx",
        index=False,
    )

    print("Enriched data quality issue files exported.")
    print("Enriched issue rows:", len(enriched_issues))
    
def main():
    data_quality_issues = []

    workbook = load_workbook(INPUT_PATH)
    validate_required_sheets(workbook, REQUIRED_SHEETS)

    standardise_column_names(workbook)
    standardise_text_values(workbook)
    standardise_status_values(workbook)
    convert_date_columns(workbook, DATE_COLUMNS_BY_SHEET)

    check_level_values(workbook, data_quality_issues)
    check_relationships(workbook, data_quality_issues)
    check_missing_evidence(workbook, data_quality_issues)
    check_stale_assessments(workbook, data_quality_issues)
    check_inactive_employee_access(workbook, data_quality_issues)
    check_high_gap_without_training_action(workbook, data_quality_issues)

    build_summary_report(workbook, data_quality_issues)
    export_data_quality_issues(data_quality_issues, OUTPUT_FOLDER)
    export_enriched_issue_report(workbook, data_quality_issues, OUTPUT_FOLDER)


if __name__ == "__main__":
    main()
