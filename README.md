# vba-repo
VB codes powering MS Office Suite tools

This repository contains codes written for tasks automation and fastening purposes, within the context of processing daily administrative and financial workloads, as well as contributing to periodical reporting, budget follow-up and account closings.

The first code (1.MASTERFILE), a VBA subprocedure, has been written with the purpose of gathering data lying in different .xls files in a single centralized file, under the form of a single continued list. The raw data in question is the registered time (hh:mm:ss), entered by every staff member, in weekly time registration files. Every row is a unique entry, where the staff member enters his/her ID, the code of the project s/he worked on, the description of the entry, and the time (hh:mm:ss).
The .xls files containing the data are grouped inside a single folder, with a .xls file for every week of the calendar year (folder meant thus to contain 52 .xls files by the end of the calendar year). 
