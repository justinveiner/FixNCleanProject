
                         ----------------------
                        ///Fix_N_Clean Readme\\\
                        ------------------------


 --------------
 Included Files
 --------------

 - Fix_N_Clean.exe
 - Volunteer_Excel_Template.xlsx
 - Client_Excel_Template.xlsx
 - Fix_N_Clean_Readme.txt


 -------------------
 Program Description
 -------------------

 This program assigns volunteers to clients and time slots for the Fix 'N' Clean
 event. The program takes volunteers and volunteer groups from an Excel sheet,
 and client locations and times from another Excel sheet. Volunteers that sign up
 in a group remain in that group, while individual volunteers are grouped together
 to ensure a minimum of three volunteers per client/time slot. The program also
 contains an automated email function that will email all volunteers their
 assignment (time, task, and location).


 --------------
 Installation
 -------------

 Copy and paste the program folder onto your computer from a USB drive.
 The copied location does not matter, just ensure that all contents of the folder
 remain in the folder and that it is in an easy to remember location. To uninstall 
 the program, delete this folder and its contents from your computer.


 ----------------
 Before you Start
 ----------------

 1. Information from the Google Forms must be put into .xlsx format (Microsoft 
    Excel). From the Google Form menu, download the data as a .csv file. Then save 
    it as a .xlsx file. Ensure that each data column matches its position in the 
    'Volunteer_Client_Template.xlsx' lists provided. Client data must also be put 
    into a .xlsx file, again matching the format of the 'Client_Excel_Template.xlsx' 
    provided. These two files may have any name.

    NOTE: The data for both files must be in 'Sheet1'.

 2. In your Gmail account settings "Allow less secure apps" must be 
    enabled to send emails using this program.


 -------------
 General Usage
 -------------

 The program will open to a main menu, giving you the option to run the volunteer
 assignment function, the email function, display hep, and quit the program. To 
 select an option, type the number to its left with no added characters. Text 
 prompts throughout the program will explain what the program is doing, and any 
 information that must be inputted. Press <Enter> to submit your input to the program.

 NOTE: For all inputs, ensure that proper capitalization is used, and avoid
 unnecessary uses of <Space>.


 --------------------
 Created Excel Sheets
 --------------------

 This program creates two Excel files, 'schedule' and 'email_failures'. 'schedule'
 is created every time the volunteer assignment function is run, and 'email_failures'
 is created every time the email function is run. To prevent the loss of data, save
 these files to a location outside of the program folder.

 'schedule' is called by specific name in the email function. As such, a .xlsx file 
 named 'schedule' must be in the same folder as the program for the email function
 to operate properly. Changes can be made to 'schedule' before emailing as long as
 the file name does not change, the file type does not change, the sheet name remains
 'Sheet1', and that all information remains in the same general format. Columns may
 be widened or shortened without consequence.


 -----------------------------
 Volunteer Assignment Function
 -----------------------------

 This function assigns volunteers to clients/time slots. The volunteers are placed
 into groups of up to four, with groups that signed up together remaining together.
 
 The program asks for the names of two Excel sheets, one containing information
 about the volunteers, and the other information about the clients. These files must
 be in a specific format, and contain specific information (see the General Usage 
 section in this file for more detailed information). The program will then assign
 volunteers to clients/time slots until all available times are full, or the list
 of available volunteers is empty. Assignment priority is in order of appearance
 in the volunteer Excel spreadsheet, and is based off of a volunteer's first and 
 second time choices. The assigned volunteers are then placed in a newly created Excel
 spreadsheet called 'schedules', that will be placed in the program folder.


 --------------
 Email Function
 --------------

 NOTE: "Allow less secure apps" in the sender's Gmail account settings must be enabled. 

 The email function makes a secure connection to a Gmail account, before sending a 
 standard message including a volunteer's assignment (tasks, time, location). The 
 function gets information for each assigned volunteer's email from 'schedule', a
 .xlsx file. This Excel spreadsheet is created by the volunteer assignment function 
 in this program. Edits can be made to this spreadsheet. However, the general 
 formatting MUST stay the same, and the sheet the data is in must be 'Sheet1'.

 The program will ask for the sender's email address and password. This information
 is only stored temporarily, and is deleted every time the program is closed and
 reopened. After the necessary inputs, the program will automatically send emails.
 Once the process is complete, the user will be returned to the main menu.

 Be aware that the sent emails may be placed into the recipients' spam folders.

 NOTE: An Excel file called 'email_failures' may be created in the same folder as
       'Fix_N_Clean.exe'. This file contains the names of any volunteers to whom
       an email was not properly sent to. These volunteers will require emails to be 
       sent manually.


 -----------
 Limitations
 -----------

 - Can only send emails from a Gmail account.
 - Each instance of the program can only assign volunteers once.


 ------------
 Known Issues
 ------------

 - Attempting to assign volunteers more than once in one instance will crash the program. To sort a 
   second time in each instance of the program, close it and reopen it.
 - Sent emails may be placed into spam folders.










