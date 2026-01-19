CT Sales Performance Automation
Overview

This project automates the end-to-end sales reporting process using Power Automate (Cloud & Desktop) and Microsoft Excel. It eliminates manual effort from data refreshes, chart updates, formatting, and report distribution, ensuring consistent and repeatable outputs every time the workflow is executed.

The automation is triggered manually and combines cloud-based orchestration with desktop-level Excel macro execution.


Tools & Technologies:

Power Automate (Cloud Flow)

Power Automate Desktop

Microsoft Excel (.xlsm with VBA macros)

Email connector



How It Works:

1. Cloud Flow (Power Automate):

Manually triggers the workflow

Retrieves the required Excel file

Sends the final report via email once processing is complete


2. Desktop Flow (Power Automate Desktop):

Launches Excel and opens the primary sales workbook

Executes a sequence of Excel macros to:

Copy updated sales data from another workbook

Refresh calculations and KPIs

Update gross revenue metrics

Regenerate and expand product-level charts

Adjust column widths for clean formatting

Saves and closes the workbook


3. Secondary Workbook Automation:

Opens a second Excel file

Runs a macro to edit the Excel sheet, edit formating and finalize downstream actions

Saves and closes Excel


Key Benefits:


Eliminates manual data refresh and formatting

Ensures consistent reporting outputs

Reduces human error

Scales easily for recurring reporting needs

Demonstrates real-world business automation skills


Use Case: 

Designed for automated sales performance reporting in a business environment, this solution is ideal for analytics, operations, and business intelligence workflows where reliability and efficiency are critical.

Notes

Macros must be enabled in Excel

File paths must match the local environment

Wait steps are included to ensure Excel completes each operation before proceeding
