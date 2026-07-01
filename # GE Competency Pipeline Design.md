# GE Competency Pipeline Design

## 1. Purpose
To clean the data set. It will normalize the format of thedata. It will also Identify data issues, relationship issues(F_K not exisiting in the main table, or vise versa),Check Business value(e.g. skill level should be within 1-5),stale data. Finally it would output an analyze ready data file.

## 2. Input
The input file is:

engineering_competency_gap_dataset_v2.xlsx

## 3. Input Standardisation
    1.Make column name in snake form. 
    2.Delete any space and replace it with "_".
    3.Making Text category in a standard format(e.g.)