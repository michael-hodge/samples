
------------------------------------------------------------------------------------------
file:
agg_spend_cross_check.xlsm

purpose:
to provide project managers a way to validate monthly aggregate spend reporting before sending to client.  upon receiving their report they drop it into the AggSpendReport tab, then update the header information on the CrossCheck tab with their client and effective dates.  those parameters are passed to a query that pulls back expense information from the database and drops the results onto a hidden tab.  program information is then brought over to the CrossCheck tab and formulas are updated based on the client selection to compare aggregate spend results to data pull.  variances are highlighted and sorted to the top for the project managers to investigate.  a GoTo Program option is added to the excel context menu to allow user to quickly navigate to the selected program in web application.


------------------------------------------------------------------------------------------
file:
cardio_speaker_tracking.xlsm

purpose:
to provide project manager with a quick method to determine if any engaged speakers are exceeding client's business rules concerning the number and type of allowable programs on a daily and annual basis.  the user updates report via a button click which connects to database and pulls detail data to programs tab.  data is then aggregated by day and by year and dropped onto the tracking tab where speakers at their engagement limit are highligted yellow and speaker exceeding their limit are highlighted red.  double clicking a row will jump to the programs tab and autofilter data to display records for the selected speaker and timeframe.


------------------------------------------------------------------------------------------
file:
mirror_audit.xlsm

purpose:
daily validation that mirror database remains in sync with production.  kicked off each morning through windows task scheduler.  exports key database tables from each system to csv files.  then brings each csv into excel and runs comparison to identify any out of sync records.  variances are saved off to another csv file.  an email is then generated notifying if whether are not any issues were found.  if there is an issue, the variance file(s) are attached to the message for review.


------------------------------------------------------------------------------------------
file:
roster_check.xlsm

purpose:
provides project manager with a more efficient way of determining updates to quarterly sales roster files from client.  on opening, the user is prompted to navigate to the previous quarter's file and then the current quarter's file.  each file is loaded into excel where they are automatically compared and any records that were changed are dropped onto a updates tab with the exact fields that were changed hightlighted.  the user can then pull up and updae those specific records in the web application instead of manually checking each one.


------------------------------------------------------------------------------------------
file:
entera_program_emails.vb

purpose:
automatic email merge stored in outlook vba project file.  each email is assigned to a customized button in the outlook ribbon.  the user selects the email they want to send and then enters the program number.  that parameter is passed to a query that pulls back the necessary information and merges it into an html generated email and auto attaches any necessary files that need to be sent with the message.


------------------------------------------------------------------------------------------
file:
program_status_rpt_format.vb

purpose:
auto format a report pulled from the web for a particular client group.  saved to personal excel workbook and run as needed.  handles work that was previously done manually.  delete columns, reformatting, splitting each project out to a seperate file, and creating a summary for each file to include in message body of email when reports are sent.


------------------------------------------------------------------------------------------
file:
agg_spend.vbs

purpose:
create monthly aggregate spend reports for clients.  client and effective dates are selected from input boxes.  the script then connects to database, reads sql file for selected client, and appends selected effective dates to query.  an excel file is created with completed and cancelled programs split onto separate tabs.  the excel file is then formatted and saved for the user.


------------------------------------------------------------------------------------------
file:
assign_points.vbs
assign_points_proc.sql

purpose:
for running a sales contest for contracted client.  sales files are received from client's sales partners and consolidated into a single file.  when run, the script picks up that file and loads it to a mysql table.  then executes a stored procedure to assign points and add notes to each record using a keywork match against the product description based on the list of systems and features the client wanted points awarded for.  it then exports the results back to an excel file which is used to determine monthly contest winners.


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
