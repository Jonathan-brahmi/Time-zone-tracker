AD Logon & Time Zone Tracker

A small tool to find the last logon and the time zone that is set on my computer.

Description: 

The Time Zone Finder Tool is a user-friendly Windows Forms application designed to help you easily gather and manage information about computer systems in your network, specifically related to last logon times, time zone (DST) status, and the ability to export data to CSV.

What Does the Script Do?
User-Friendly Interface: The tool features a modern, dark-themed user interface that mimics the look of popular systems like Ubuntu. It includes a Form (window) with buttons, text fields, and labels for a smooth and intuitive user experience.
Search Computers in Active Directory (AD):
When you click the "Start the Process" button, the tool begins scanning your Active Directory (AD) to find all computers and retrieve important information, including the last logon date and whether the system is currently in Daylight Saving Time (DST).
For each computer, the following details are gathered:
Computer Name
Last Logon Date
DST Status (whether it’s in Summer or Winter time)
Local Time of the query
Find Specific Computer:
You can search for a specific computer by typing its name into the search field and clicking "Find Computer". The system will display the status and local time for the computer you searched for, allowing you to quickly pinpoint details.
Export Data to CSV:
After running the scan, you can export the results (including the last logon data and time zone information) into a CSV file by clicking the "Export to CSV" button.
The CSV file will include all the details for the computers found during the scan, and you can easily save it to your preferred location.
Real-Time Updates:
The application provides real-time feedback through a status label that informs you of the process (e.g., when the scan starts, when data is ready to save, or if no computers were found).
It also uses Out-GridView to display the results in a grid view for better readability.
Key Features:
Modern, Dark UI: The tool comes with a sleek and modern interface, including clear buttons and a clean design.
Active Directory Integration: Easily find computers in your network, display their last logon times, and check the DST status.
Customizable Search: Search for specific computers by name with ease.
Data Export: Export computer details to CSV for further analysis or reporting.
Real-Time Status Updates: Receive feedback throughout the process with real-time status changes.
Requirements:
Windows with PowerShell enabled.
Active Directory access to query computer details.
This tool is ideal for IT administrators, network managers, or anyone who needs to quickly gather computer logon and time zone information across their systems.




Install instructions
How to Download and Run the AD Logon & Time Zone Tracker
Follow these steps to download and run the AD Logon & Time Zone TrackerPowerShell script on your Windows machine:

Step 1: Download the Script
Go to the Time Zone Finder project page on itch.io. (Replace the # with the actual URL of your itch.io page.)
On the project page, click the Download button to download the script.
The file will be downloaded as a .zip or .ps1 file.
If the file is in a .zip archive, extract it to a folder on your computer.
Step 2: Enable PowerShell Scripts (If Necessary)
To run PowerShell scripts, you may need to allow script execution on your system. Here’s how to do it:

Press Win + X and select Windows PowerShell (Admin) to open PowerShell as an administrator.
Run the following command to enable script execution:
powershell
Copy code
Set-ExecutionPolicy RemoteSigned 
When prompted, press Y and then Enter to confirm the change.
This will allow you to run local scripts safely while still protecting your system from potentially harmful scripts downloaded from the internet.

Step 3: Run the Script
Open PowerShell on your computer. You can do this by searching for "PowerShell" in the Start menu and selecting Windows PowerShell.
Navigate to the folder where the script is located. For example, if the script is in the "Downloads" folder:
powershell
Copy code
cd "C:\Users\YourUsername\Downloads" 
To run the script, enter the following command:
powershell
Copy code
.\TimeZoneFinder.ps1 
If you extracted the script from a .zip file, navigate to the folder where the script was extracted.

Step 4: Follow the On-Screen Instructions
Once the script starts, a window will appear with buttons and prompts. Here's what to expect:

Start the Process: This button will begin scanning for computers in Active Directory.
Find Computer: Allows you to search for a specific computer by name.
Export to CSV: Saves the results of the scan into a CSV file.
The results will show the last logon time, time zone information, and other details related to the computers found in Active Directory.

Troubleshooting
If you encounter any issues running the script, make sure you have the required permissions (running PowerShell as an administrator) and that script execution is enabled as described in Step 2.
If the script encounters issues with Active Directory, make sure you have the appropriate rights to access Active Directory information and that the necessary PowerShell modules (such as ActiveDirectory) are installed.
