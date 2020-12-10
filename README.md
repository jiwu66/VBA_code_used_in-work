# VBA_code_used_in-work
Some VBA code I used during coop term as Data and Application Support

Sample input and expected output files are in included.

The goal was to extract data from the source sheet and enter them into the target sheet. The code will compare their Operand ID(Unique Key), if the Operand ID matches the data in the source sheet will be entered. If the filed in the target sheet is already taken(When there are more than one file for one people), the code will create a new row in the target sheet and enter the data there.
The code will mark "1" and the end of each row in the source sheet if that row has been entered, this will make it easier to check if the code is doing what we wanted. We can simply look at the "File Uploaded Column" to debug our code.
