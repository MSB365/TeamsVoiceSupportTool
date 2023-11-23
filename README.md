# TeamsVoiceSupportTool
This PowerShell script is designed to assist Microsoft Teams administrators in efficiently handling bulk assignments of phone numbers and user policies. It streamlines the process of assigning phone numbers and user policies to multiple users within Microsoft Teams.

**Features**

Bulk Phone Number Assignment: Assign phone numbers to multiple users at once.
User Policy Assignment: Assign user policies to a group of users for streamlined management.
Customization: Easily customize the script to suit specific organizational requirements.



**Prerequisites**

PowerShell Version: Ensure PowerShell 5.1 or later is installed.
Microsoft Teams Admin Center Access: Administrator access to Microsoft Teams Admin Center is required.
CSV File: Prepare a CSV file containing necessary user information (e.g., user IDs, phone numbers, policy assignments).



**Usage**

Download: Clone or download the script to your local machine.
Prepare CSV: Create a CSV file following the provided template with required user details.
Modify Script: Update script variables such as paths, policies, and API keys as per your environment.
Execute Script: Run the PowerShell script in an elevated PowerShell session.



**CSV File Format**

Depending of the use area, two CSV formats are required:
Phone Number Management: "UserPrincipalName","DisplayName","TelephoneNumber","CallingPolicy"
Policy Management: "UserPrincipalName","TeamsPolicy"



**Important Notes**

Backup Data: Always maintain a backup of user information before bulk assignments.
Permissions: Ensure appropriate permissions for the script execution.
Testing: Test the script in a non-production environment before applying changes in a live environment.



**Additional information**

This script is based on https://github.com/MSB365/PhoneNumberAssignment-V2.0 (Documentation at the GitHub Link and/or https://www.msb365.blog/ )
New in this Script, compared to the PhoneNumberAssignment Script, is the BULK User policy Management.
