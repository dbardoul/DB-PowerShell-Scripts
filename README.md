# Introduction:
This GUI was designed to make launching my scripts easier and faster. Read the description above each button to know how the script will perform. Upon button click, a seperate PowerShell window will appear and begin the script. The window must be manually closed by the user when finished. This gives the user the ability to fully read their input and output before ending. After the PowerShell window is closed, the user can click another button or close the GUI.

##### Prerequisites to run GUI
- Powershell v5
- Local administrator on machine

##### Prerequisites to run every script
- Microsoft Excel
- Remote Server Administration Tools


# Main.ps1:
Main.ps1 is the PowerShell script that contains the GUI. It consists of 8 labeled buttons and a short description of each script.


# Copy-Files.ps1:
Copy-Files.ps1 is the PowerShell script that runs when the Copy-Files button is clicked. Its purpose is to copy one or more files to a group of hostnames. This is useful for the planning stage of a practice aquisition or expansion project. An example of its use is to copy one or more internet explorer or chrome shortcuts to a list of PCs being prepped for a project.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Prompts for the source file(s).
3. Prompts for a selection between the following destinations:
   - C:\Users\Public\Desktop
   - C:\Temp
4. Copies file(s) to the selected destination on each hostname.


# Hostname-Report.ps1:
Hostname-Report.ps1 is the PowerShell script that runs when the Hostname-Report button is clicked. Its purpose is to output an excel spreadsheet containing AD information for a group of hostnames. This is useful for the post-imaging stage of a PC. An example of its use is to check if all the imaged PCs were added to the right OU and AD groups.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Prompts for the location to store the excel spreadsheet.
3. Generates the spreadsheet. Contents listed below:
   - Hostname
   - OU
   - Ping status
   - AD groups
  
  
# Add-ADGroup.ps1:
Add-ADGroup.ps1 is the PowerShell script that runs when the Add-ADGroup button is clicked. Its purpose is to add a group of hostnames to a selected AD group. This is useful for the post-imaging stage of a PC and for administrative tasks. An example of its use is for Linh, Alex, or I to bulk add hostnames to the Office_2016_installer group.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Prompts for AD credentials.
3. Prompts for a selection between the following AD groups:
   - Office_2016_installer
   - Greenway-Computer
   - MBAM_Installer
   - JavaRuleSet_Bypass_Computers
4. Adds the list of hostnames to the selected AD group.


# Compare-KBs.ps1:
Compare-KBs.ps1 is the PowerShell script that runs when the Compare-KBs button is clicked. Its purpose is to compare the installed KBs on a list of hostnames against a list of KBs to search for matches. This is useful for when TORK releases a statement identifying a problem KB. A previous example of its use was when there were a list of KBs that caused problems for Greenway and needed to be uninstalled. I used the script to search for the KBs on each hostname when a ticket came in.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Prompts for a text or csv file that contains a list of KBs.
3. Prompts for the location to store the excel spreadsheet.
4. Generates the spreadsheet. Contents listed below:
   - Hostname
   - List of KBs
   - If there is a match, the cell is colored green. If not, it is colored red.
  
  
# Move-ADObject.ps1:
Move-ADObject.ps1 is the PowerShell script that runs when the Move-ADObject button is clicked. Its purpose is to move a list of hostnames to a selected OU. This is useful for the post-imaging stage of a PC and the planning stage of a practice aquisition or expansion project. An example of its use is for bulk moving a list of hostnames to the proper OU post-imaging.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Prompts for AD credentials.
3. Prompts for the OU selection.
4. Moves the list of hostnames to the proper OU.


# Remove-ADComputer.ps1:
Remove-ADComputer.ps1 is the PowerShell script that runs when the Remove-ADComputer button is clicked. Its purpose is to remove a list of hostnames from AD. This is useful for the pre-imaging stage of an existing PC. An example of its use is for removing a hostname from AD so that it can be imaged.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Prompts for AD credentials.
3. Removes the list of hostnames from AD.


# Rename-Computer.ps1:
Rename-Computer.ps1 is the PowerShell script that runs when the Rename-Computer button is clicked. Its purpose is to rename a list of hostnames. This is useful for the planning stage of a practice aquisition or expansion project and prepping a new PC for a ticket. An example of its use is for bulk renaming a group of pre-imaged stockroom PCs that already have names.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Prompts for AD credentials.
3. Prompts for user input for the following fields:
   - Beggining hostname acronym (e.g. MIS, PP, RAD)
   - Ending hostname acronym (e.g. GR, NSA, NSF)
4. Renames the list of hostnames using the begging acronym, last 5 of the serial, and ending acronym.


# WiFi-Profiles.ps1:
WiFi-Profiles.ps1 is the PowerShell script that runs when the WiFi-Profiles button is clicked. Its purpose is to add NSHData5 and NSHDataRS to a list of hostnames. This is useful for the post-stage of imaging a laptop.

##### Performance
1. Prompts for a text or csv file that contains a list of hostnames.
2. Adds NSHData5 and NSHDataRS to each hostname.
