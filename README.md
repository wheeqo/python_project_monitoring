# Python EPC Project Monitoring from Multiple Excel Schedule with 3-Stage Submission
## How I Use Data Analyst Skill to Create Project Monitoring with Python

This case study is my experience as an employee in EPC (Engineering, Procurement, Construction) Company. The company has several multi-disciplinary projects, and each project schedule is managed with Excel file with some projects have 2-stage submission and some have 3-stage submission.

It is impractical to check Excel file one by one which deliverable item that must be submitted, which one is due to, or how much the project progress is. To make things worse, each Excel project schedule have different format. Therefore, I took initiative combining all excel project schedule then querying & filtering on Python then give output directly to IDE software (I am using VS Code). The output report to be displayed are: 
-	Table of late deliverable item list,
-	Plotly diagram showing how many late deliverable items each discipline and each project, 
-	Plotly diagram showing how much each project progress (planned vs. actual), and
-	Plotly diagram showing S-Curve for each project.

Then the output report above will be exported as single HTML, therefore can be easily printed to PDF from browser.
