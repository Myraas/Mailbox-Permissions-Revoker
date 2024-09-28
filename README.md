## Mailbox-Permissions-Revoker

**Description:**

MailboxPermissionsRevoker.ps1 is a PowerShell script for administrators to find and remove mailbox permissions (FullAccess, SendAs, and SendOnBehalf) assigned to a specific user in Exchange Online. The script scans for these permissions throughout the tenant and provides options to remove them selectively or all at once. It can also import permissions from a CSV file if needed.

**Features:**

- Finds and lists all FullAccess, SendAs, and SendOnBehalf permissions assigned to a specific user.
- Provides the option to remove permissions selectively or all at once.
- Allows importing cached permissions from a CSV file for faster operation.

**Caching Permissions:**

To cache mailbox permissions for faster processing, you can run the AdminDroid "GetMailboxPermission.ps1" (version 3.0) script in the same directory as this script. This will generate a CSV file containing the permissions that can be used by MailboxPermissionsRevoker.ps1.

You can find the AdminDroid script here:
[GetMailboxPermission.ps1](https://github.com/admindroid-community/powershell-scripts/blob/master/Office%20365%20Mailbox%20Permissions%20Report/GetMailboxPermission.ps1)

Note: This script is created and maintained by AdminDroid. Please refer to their repository for updates and support. Special thanks to AdminDroid for providing a tool that complements this script’s functionality.


## Disclaimer
Please note that while I have taken care to ensure the script works correctly, I am not responsible for any damage or issues that may arise from its use. Use this script at your own risk.

## License
This project is licensed under the terms of the GNU General Public License v3.0.
