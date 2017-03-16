------------------------------------------------------------------------------------------
file:
Provider DBQA.sas

purpose:
this was initially a webi report in business objects.  the user requested additional information be added to it that is housed in a different source system.  there was no way to tie the two universes together in business objects so i re-wrote the report using sas.  the program pulls down the necessary data from each system and joins it together then exports to a data tab on an excel file that is accesses via brower in sharepoint.

------------------------------------------------------------------------------------------
file:
Weekly Provider Sample.sas

purpose:
request was to fulfill an audit requirement that each week, any data changes have to be reviewed for for accuracy for 50 provider records.  the sas program connects to the database and pulls in any provider record that was created or edited in the last seven days and selects a sample of 50, saves them to excle, and sends them out as an email attachment. 

------------------------------------------------------------------------------------------
file:
N00300 Verify ClmTyp and Spclty mapping.sas

purpose:
one of a couple hundred data quality queries that run each night to find problem records for the configuration team to review and correct.  this one locates specific combinations of specialty and claim type that should not exist and saves them to excel.

------------------------------------------------------------------------------------------
file:
F00413 PROVs Active in both TIER1 and TIER2.sas

purpse:
another data quality query.  this one identifies any provider who at any time has been linked to both TIER1 and TIER2 networks at the same time.

------------------------------------------------------------------------------------------
file:
query_dashboard.xlsm
query_dashboard.mht

purpose:
to provide management with information about error fallout from nightly data quality queries.  the report buckets the information by category and development status.  other sections display high level error counts, trending, and a list of the current high volume queries that need to be worked first. vba code retrieves data from different sources and drops to data tabs, performs some formatting, then publishes report as .mht file to a sharepoint page.  once is sharepoint it's largely static, except for file links to each high volume output file and a filter on the page to allow navigation to previous day's reports.    
------------------------------------------------------------------------------------------
file:
query_health.xlsm

purpose:
a high level daily overview of error counts from the nightly data quality queries bucketed by system.  vba code retrieves data from different sources and drops to data tabs.   

------------------------------------------------------------------------------------------
file:
password_exp.xlsm

purpose:
to notify users of upcoming database password expirations so they can be updated before multiple overnight jobs attempt to login with bad credentials, locking the account.  the program runs daily and queries each database system that the group supports for password expiration.  a reminder email is sent once a week until a password is within 14 days of expiring.  then a warning email is sent out daily to the group until the password is updated.  written in vba instead of vbs due to company policy concerning script files.

------------------------------------------------------------------------------------------
file:
mm_db_test.accdb

purpose:
proof of concept for a project.  the requesting group ran reports from an access database using close to 100 imported spreadsheets as tables.  they wanted to have a way to verify that all tables were up to date prior to running reports.  the spreadsheets were maintained and updated intermittently by different groups.  they couldn't link directly to the spreadsheets without locking them from being edited and they didn't want to have a scheduled nightly refresh of all of the tables.  so this form gave them a quick way see if any files had been modified since the last table refresh date and to import only those tables with a button click.  vba code runs on form load to get the file modified date of each file that is the basis for a table and an indicator shows whether the tables need to be refreshed.

------------------------------------------------------------------------------------------

file:
agg_spend_cross_check.xlsm

purpose:
to provide project managers a way to validate monthly aggregate spend reporting before sending to client.  upon receiving their report they drop it into the AggSpendReport tab, then update the header information on the CrossCheck tab with their client and effective dates.  those parameters are passed to a query that pulls back expense information from the database and drops the results onto a hidden tab.  program information is then brought over to the CrossCheck tab and formulas are updated based on the client selection to compare aggregate spend report results to data pull.  variances are highlighted and sorted to the top for the project managers to investigate.  a GoTo Program option is added to the excel context menu to allow user to quickly navigate to the selected program in the web application.

------------------------------------------------------------------------------------------
file:
cardio_speaker_tracking.xlsm

purpose:
to provide project manager with a quick method to determine if any engaged speakers are exceeding the client's business rules concerning the number and type of allowable programs on a daily and annual basis.  the user updates report via a button click which connects to the database and pulls detail data to programs tab.  data is then aggregated by day and by year and dropped onto the tracking tab where speakers at their engagement limit are highligted yellow and speakers exceeding their limit are highlighted red.  double clicking a row will jump to the programs tab and autofilter the data to display records for the selected speaker and timeframe.

------------------------------------------------------------------------------------------
file:
mirror_audit.xlsm

purpose:
daily validation that mirror database remains in sync with production.  kicked off each morning through windows task scheduler.  exports key database tables from each system to csv files.  then brings each csv into excel and runs comparison to identify any out of sync records.  any issues are saved off to a variance file.  an email is then generated notifying whether are not any issues were found.  if there is an issue, the variance file(s) are attached to the message for review.

------------------------------------------------------------------------------------------
file:
roster_check.xlsm

purpose:
provides project manager with a more efficient way of determining updates to quarterly sales roster files from the client.  on opening, the user is prompted to navigate to the previous quarter's file and then the current quarter's file.  each file is loaded into excel where they are automatically compared and any records that were changed are dropped onto an updates tab with the exact fields that were changed hightlighted.  the user can then pull up and update those specific records in the web application instead of manually checking each one.

------------------------------------------------------------------------------------------
file:
entera_program_emails.vb

purpose:
automatic email merge stored in outlook vba project file.  each email is assigned to a custom button in the outlook ribbon.  the user selects the email they want to send and then enters the program number.  that parameter is passed to a query that pulls back the necessary information and merges it into an html formatted email and auto attaches any necessary files that need to be sent with the message.

------------------------------------------------------------------------------------------
file:
program_status_rpt_format.vb

purpose:
auto format a report pulled from the web for a particular client group.  saved to personal excel workbook and run as needed.  handles work that was previously done manually: deleting columns, reformatting, splitting each project out to a seperate file, and creating a summary for each file to include in message body of email when reports are sent.

------------------------------------------------------------------------------------------
file:
agg_spend.vbs

purpose:
create monthly aggregate spend reports for clients.  client and effective dates are selected from input boxes.  the script then connects to the database, reads sql file for selected client, and appends selected effective dates to query.  an excel file is created with completed and cancelled programs split onto separate tabs.  the excel file is then formatted and saved for the user.

------------------------------------------------------------------------------------------
file:
assign_points.vbs
assign_points_proc.sql

purpose:
for running a sales contest for the contracted client.  sales files are received from client's sales partners and consolidated into a single file.  when run, the script picks up that file and loads it to a mysql table.  then executes a stored procedure to assign points and add notes to each record using a keyword match against the product description based on the list of systems and features the client provided.  it then exports the results back to an excel file which is used to determine monthly contest winners.

------------------------------------------------------------------------------------------
file:
entera_final_headcount.vbs


purpose:
automated email script.  kicks off each morning through windows task scheduler.  connects to the database and checks if the client has any dinner programs scheduled for one week from that day.  if there are any, it loops through each program to generate a reminder email that gets pushed to the rep responsible for the program to provide final headcount numbers.


------------------------------------------------------------------------------------------
file:
insys_program_approval_email.vbs


purpose:
automated email script.  kicks off each morning through windows task scheduler.  connects to the database and checks if the client has any programs in pending approval status.  if there are, then an email is generated listing the program information and a reminder for the managers to log into the sales portal to approve or decline each one.

------------------------------------------------------------------------------------------
file:
after_hours_report.sql

purpose:
provide staff working evenings with program information needed if sales rep calls with questions.  feeds into web report.

------------------------------------------------------------------------------------------
file:
hangman.xlsm

purpose:
it's hangman.  in excel.  because why not?

------------------------------------------------------------------------------------------
file:
ssrs_test.pdf (via SSRS)

purpose:
a couple of reports originally developed using iReport.  re-created in SSRS as a test.

------------------------------------------------------------------------------------------
file:
grifols_sign_in_sheet.pdf (via iReport)

purpose:
sign in sheet provided to pharmaceutical sales reps to collect information of health care providers attending speaker programs.

------------------------------------------------------------------------------------------
file:
pernix_sign_in_sheet.pdf (via iReport)

purpose:
sign in sheet provided to pharmaceutical sales reps to collect information of health care providers attending speaker programs.

------------------------------------------------------------------------------------------
file:
insys_invitation.pdf (via iReport)

purpose:
speaker program invitation provided to pharmaceutical sales rep to distribute to target attendees.

------------------------------------------------------------------------------------------
file:
retrophin_cam_evaluation.pdf (via iReport)

purpose:
speaker program evaluation form provided to pharmaceutical sales rep.

------------------------------------------------------------------------------------------
file:
united_therapeutics_program_checklist.pdf (via iReport)

purpose:
used internally by project coordinators to track the necessary steps for each speaker program they manage.

------------------------------------------------------------------------------------------
file:
insys_budget_vs_spend.xlsm

purpose:
the client's budget for speaker programs is divided between close to 300 seperate territories.  the project manager would receive weekly requests to remove money from some territories and re-allocate among others.  the decision on where to re-allocate is made by the project manager based on each territories current budget and spend.  he previously accomplished this by running and exporting two seperate reports, then manually combining them with a budget tracking spreadsheet and filtering everything to see the information he needed.  this updated process automatically pulls in all of the necessary information from the database and budget tracking spreadsheet and merges it onto one tab.  since the the client's budget is done on a quarterly basis, a control was added to allow the project manager to automatically switch quarter views.  an additional feature was added so when he clicks in the grid, it will automatically highlight all of the data for that territory to make it easier to zero in on the information he needs.

------------------------------------------------------------------------------------------
file:
program_metric_tracking.xlsx (via iReport)

purpose:
the program director requested a metric report to track key activities performed during the life of each program for client review and internal benchmarking purposes.  some of this information is stored in a table of unserialzed log data and isn't optimized for reporting.  so a stored procedure was created to retrieve the necessary metric information and dump it into a smaller, indexed table that then feeds into the metric report.
