# vba-repo
<u><b>VB codes powering MS Office Suite tools</b></u>

This repository contains codes written for tasks automation and fastening purposes, within the context of processing daily administrative and financial workloads, as well as contributing to periodical reporting, budget follow-up and account closings.

<u><b>Deadline Notification</b></u>

This is referring a VBA module containing and combining a sub routine and a user-defined function.
The function is based on a date-type parameter and checks out whether a deadline is nearing, via the use of the built-in VBA function DATEADD, and returns a boolean type result; with True result in case the tested value falls within 30 days from today's date (date on which the code is run). 
The subroutine starts by declaring relevant variables and refers to a a specific column inside a specific table, located in a specific worksheet of a workbook. Via a For Each - Next loop, the value of every cell (thus, containing deadline date) is tested through the function which I defined; and a notification email is dispatched for any deadline which is nearing, in line with the period having been fixed in the function. The email contains additional variable data, such as the description of the product whose deadline is approaching, its supplier, along with the deadline date in question. Along with sending emails, the sub also adds to the table (in a column dedicated for that end) the information that the email notification has been sent. If an email notification has already been sent for a certain deadline, a second email is not forwarded; the If statement has been typed by taking this criteria/standard into account and to avoid any notification duplicate.

<u><b>Masterfile</b></u>

This code, a VBA sub procedure, has been written with the purpose of gathering data, lying originally in different .xls files, within in a single centralized masterfile under the form of a single continued list. The raw data in question is the registered time (hh:mm:ss), as entered by every staff member, in weekly time registration files along with other information. Every row is a unique entry, where the staff member enters own ID, the code of the task s/he worked on, the description of the entry, and the time (hh:mm:ss) which was spent.
The .xls files containing the data are located inside a single folder, with one .xls file being assigned for every week of the calendar year (hence, the folder being meant to contain 52 .xls files by the end of the calendar year). 
The VBA macro code I wrote contains Do Until, For Next, Do While and For Each Next loops. An integer-type variable is also added and incremented from 1 onwards, inside Do While Loop which is based on opening every single .xls file inside the Directory.
The code works succesfully, first creating a new workbook on the Desktop, with three relevant sheets being created therein (after the names and structure of the files containing the data to be transferred), all the while deleting the already-existing default sheets at the creation of the workbook. Thereafter, the code copy-pastes the header parts (standard to every weekly .xls file) from one of the files in the directory, and this, for each respective sheet in the masterfile. That way, all 3 sheets in the Masterfile are ready for the data gathering. We add an extra column, which shall contain (for every single imported data row) the reference of the week through the incremented integer variable; being incremented according to the reference of the week for which data is being transferred (1 to 52). Once the collection of data in the centralized file is complete, the macro adapts the format of the column containing the time to a proper time-number format in Ms Excel, the macro runtime is then completed and this is announced through a message box. The data centralized in the masterfile is then ready for further analysis and registration.

<b>Email_Automation</b>

Based on a For Each Next Loop integrating an inner condition, this code generates automated emails via the outlook emailing app, using a text template as message body, and customizing each email to be sent on the basis of data retrieved from an .xls file's table; with one email being sent for each row, insofar as the conditions which were set out are being met. Each email is dispatched based on variables referencing the key information (name and email address of the addressee, and other information) which are retrieved from the relevant columns of the active row. 
The code starts with the declaration of object variables to activate and refers to relevant applications and files, as well as to the table from which most of the variable data shall be retrieved. Later, the email is composed with the referencing of variables; the body of the message is in html format, so that customized, ad hoc adaptations to the format can be performed easily. The outlook signature of the user is also included in the emails which are going to be forwarded. The running of the code is completed with the dispatching of the emails for all rows.   

<u><b>Updating Tasks List</b></u>

This macro is embedded to a .xls file being utilized as an agenda, encompassing list of tasks to do, organised under the form of table. The table is made up of several columns identifying each task (one row/line being a task), providing various information and descriptions, with the last column (column I) indicating the current status. The status can take several forms, "Open", "Pending", "On Hold", as well, "Completed" once the task is fulfilled.
The macro aims at, once being run, hiding rows whose status have been changed to "Completed" and adding a certain number of new blank rows at the bottom of the table (to be available for further utilisation). The code starts with a Do While loop hiding rows containing tasks being "completed", then a For Next loop is being run to add new blank rows based on user InputBox.
