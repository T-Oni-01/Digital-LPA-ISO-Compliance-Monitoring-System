# Digital-LPA-ISO-Compliance-Monitoring-System
A constraint-based LPA scheduler that enforces shift compatibility, balances audit workload, prevents repetitive pairings, and automatically adapts to staffing changes. Each auditor is assigned exactly two LPAs per month, and every assignment is validated against ISO audit rules.


***System Overview:
Tab	Function
Tab 1	Setup
Tab 2	Smart LPA Scheduler (constraint-aware, adaptive)
Tab 3	Month-by-Month LPA Completion Tracker
Tab 4	ISO & Corrective Action Aging & Overdue Flags

Tab 1: Setup
•	User Inputs:
o	Add / remove employees
o	Assign role & shift
o	Add / remove plant sections
o	Mark employees active/inactive

Tab 2: SMART & ADAPTIVE LPA SCHEDULER
•	Functional Requirements:
o	Inputs
o	Personnel
o	First name
o	Last name
o	Role (Engineer / Supervisor / Manager)
o	Shift (1st / 2nd / 3rd)
o	Plant Sections
o	Dynamically add/remove (311, 341, 361, etc.)
o	Month to schedule
•	Scheduling Rules (Constraints)
o	LPA performed in pairs of 2 (3 only if needed)
o	Shift compatibility enforced
o	Avoid repeating the same pair within 4 months
o	Engineers / supervisors / managers allowed
o	 Automatically adapts to hires / removals
•	Shift Compatibility Logic
Target Shift	Allowed Auditors	Forbidden Pairs
1st 	1st + 2nd or 1st + 3rd or 1st + 1st 	2nd + 3rd 
2nd 	2nd + 1st or 2nd + 3rd or 2nd + 2nd 	1st + 3rd 
3rd 	3rd + 1st or 3rd + 2nd or 3rd + 3rd 	1st + 2nd 
•	User-Driven Inputs

o	A. Plant Sections
	Add / remove sections dynamically
	Example: 311 – Coil Winding, 341 – Internal Assembly, etc.

o	B. Employees
	Each employee has:
	First Name
	Last Name
	Role (Engineer / Supervisor / Manager)
	Shift (1st / 2nd / 3rd)
	Active / Inactive status (for hires & removals)

•	Monthly LPA Requirements
o	 Each employee must perform exactly 2 LPAs per month
o	 LPAs are performed in pairs of 2
o	3-person LPAs allowed only as a fallback (edge case)
o	LPAs are assigned to:
	A plant section
	A target shift (1st, 2nd, or 3rd)

Tab 3 — LPA Completion Tracker
•	Completed vs missed LPAs
•	Completion rate by:
o	Section
o	Shift
o	Role

Tab 4 — ISO & Corrective Actions
•	Nonconformities
•	Due dates
•	Automatic overdue flagging
•	Aging analysis

<img width="898" height="538" alt="image" src="https://github.com/user-attachments/assets/5a9659a0-24a2-421e-a239-a174fb327c59" />

<img width="900" height="511" alt="image" src="https://github.com/user-attachments/assets/6f3e0600-3298-4596-a03b-427be9f89697" />

<img width="959" height="507" alt="image" src="https://github.com/user-attachments/assets/886714d6-598f-482a-b6d4-eda891272eba" />

<img width="901" height="471" alt="image" src="https://github.com/user-attachments/assets/6bd66bde-4ad7-4e03-a0e4-025eabe906f3" />

<img width="897" height="477" alt="image" src="https://github.com/user-attachments/assets/d8a1c9e1-0850-42da-8749-370e0b0c10e4" />



