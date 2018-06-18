# mRLCDataTool

This spreadsheet receives raw test data from a user, scores and codes the input, and generates a variety of reports for different audiences. 

2.21 - June 14, 2018
Fixed outcome entry for 2018 grade 7 Shape and Space.

2.23 - June 16, 2018
version reverted to 2.2 and above fix applied to address file rename crash.

2.3b - June 18, 2018
Crashing still occuring. Took the following steps:
- separated the userform combobox population routines (Year, Grade, Division) into new subs and called them after the form was loaded
- tightened up some of the loop ranges used when populating the comboboxes to speed things up a bit
- removed the histogram sheets - will have to re-create later
- discovered that the crash would not happen if the VBA Editor was open. Inserted a line to open the VBA Editor at workbook launch. This seems to have solved the crash problem, but introduces code security issues.
