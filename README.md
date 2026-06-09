# Absolute-macros.in-excel
📊 Excel VBA Absolute Macro Creation – README
Overview

This project demonstrates how to create and automate an Absolute Reference Macro in Microsoft Excel using VBA (Visual Basic for Applications).

The macro generates a summary table from employee data by grouping records based on the Color field and calculating:

Count of records for each color
Total Bonus for each color
Total Salary for each color
Automated formatting and table styling

Since the macro uses absolute cell references, it always writes output to the same worksheet locations regardless of the currently selected cell.

Features

✅ Automated summary table creation

✅ Uses Absolute References

✅ Calculates record count by category

✅ Calculates total salary by category

✅ Calculates total bonus by category

✅ Applies formatting automatically

✅ Beginner-friendly VBA project

Dataset Structure
ID	Name	Salary	Bonus	Average	Color
1016	Dustin Henderson	626562	1000	313781	Red
1017	Lucas Sinclair	96526	1000	48763	Blue
1018	Will Byers	566458	1000	283729	Green
Technologies Used
Microsoft Excel
VBA (Visual Basic for Applications)
Excel Developer Tools
How to Run
Step 1: Enable Developer Tab
Open Excel
Go to File → Options
Select Customize Ribbon
Enable Developer
Click OK
Step 2: Open VBA Editor

Press:

ALT + F11
Step 3: Insert Module
In VBA Editor
Click:
Insert → Module
Paste the macro code.
Step 4: Save Workbook

Save as:

Excel Macro-Enabled Workbook (*.xlsm)
Step 5: Run Macro

Press:

ALT + F8

Select:

absolute_stranger

Click Run

Absolute Reference Concept

An Absolute Macro records actions with fixed cell references.

Example:

Range("H11").Value = "Color"
Range("I11").Value = "Count"
Range("J11").Value = "Total Bonus"

No matter where the cursor is located, the values are always written to:

H11
I11
J11
Sample Output
Salary Summary
Color	Count	Total Salary
Red	3	749680
Blue	3	1007019
Green	3	971740
Bonus Summary
Color	Count	Total Bonus
Red	3	3000
Blue	3	3000
Green	3	3000
Learning Outcomes

By completing this project, you will learn:

VBA Basics
Macro Recording
Absolute References
Excel Automation
Range Selection
Cell Formatting
Data Aggregation
Report Generation
Project Structure
Pivot-Tablesdashboard1.xlsm
│
├── Data Sheet
├── Summary Tables
├── VBA Module
│   └── absolute_stranger()
└── Dashboard Sheets
Future Improvements
Convert to Dynamic Macro
Add User Forms
Generate Charts Automatically
Create Dashboard Visualizations
Export Reports to PDF
Add Error Handling
Author

Thanuja Tirumalasetty

Aspiring Data Analyst | Excel | VBA | SQL | Python | Power BI
