# vba-repo
<u><b>VB codes powering MS Office Suite tools</b></u>

This repository contains codes written for tasks automation and fastening purposes, within the context of processing daily administrative and financial workloads, as well as contributing to periodical reporting, budget follow-up and account closings.

<u><b>Masterfile</b></u>

The first code, a VBA subprocedure, named Masterfile, has been written with the purpose of gathering data, lying originally in different .xls files, within in a single centralized masterfile under the form of a single continued list. The raw data in question is the registered time (hh:mm:ss), entered by every staff member, in weekly time registration files along with other information. Every row is a unique entry, where the staff member enters his/her ID, the code of the task s/he worked on, the description of the entry, and the time (hh:mm:ss) which was spent.
The .xls files containing the data are located inside a single folder, with one .xls file being assigned for every week of the calendar year (hence, the folder being meant to contain 52 .xls files by the end of the calendar year). 
The VBA macro code I wrote is a nested Do While Loop; with the first loop based on the directory and the .xls files being saved therein, and the second loop being based on cells within each cell range inside the concerned sheet of the specific weekly time registration file (having been activated in the first loop). An integer type variable is being incremented from 1 onwards, inside the first loop.
The code works succesfully, by opening every xls file in the folder, activating the relevant sheet before copying every row containing time registration data of the week in question, then returning to the Masterfile where the data is centralized, with every copied row being pasted to the last plus one cell containing data. Thanks to the variable being put inside the first loop, we could track which week (1 to 52) of the year any given entry pertains to. This information is being added to the masterfile, with the fifth column (column e) of every row being annotated with the week reference of the time registration entry. Once the collection of data is the centralized file is complete, the macro adapts the format of the column containing the time to a proper time-number format in Ms Excel, the macro is being fully run. The data centralized in the masterfile is then ready for further analysis and registration.

<b>Email_Automation</b>
