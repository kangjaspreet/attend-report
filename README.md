# Upwork Excel Data Extraction Job Description
I have a folder with 400+ .xlsx files that needs to be merged together into one .csv file. In this GitHub repo, I have provided six example files. Each file **name** follow the format of "Test_YYYY-MM-DD_[GroupID].xlsx". For example, "Test_2018-05-11_A.xlsx" is the report for Group A on the date of 2018-05-11, and "Test_2018-05-20_A.xlsx" is the report for Group A on 2018-05-20, etc.

The issue is that the exact number of rows and their layouts differ between files. This can be seen when comparing between groups A-B-C on the same date as well as the same group on different dates. For example, if you look at Group C's report, you will find that [AC-No] 2378 (in the merged column E-G in the worksheet) is present in 2018-05-11 but absent in 2018-05-20.

The result that I am looking for can be seen in the example file **"Test_Merged.xlsx"** This file removes all unused info and keeps only the relevant columns that can be used for later analyses.

# The Job:
Write a **Python** or **R** script that extracts the information shown in **"Test_Merged.xlsx"** from the six test .xlsx files and combine them together into one **.csv** file using the same format as **"Test_Merged.xlsx"** 
