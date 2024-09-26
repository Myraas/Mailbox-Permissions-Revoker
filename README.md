## Mailbox-Permissions-Revoker

**Description:**

MailboxPermissionsRevoker.ps1 is a PowerShell script designed to help administrators find and manage mailbox delegations assigned to a specific user in Exchange Online. The script scans for FullAccess, SendAs, and SendOnBehalf permissions that have been delegated to the target user. It provides options to selectively or fully remove these delegations. This script is a handy tool for admins managing mailbox permissions cleanup and maintaining security and compliance in Exchange environments.

**Features:**

- Finds and lists all FullAccess, SendAs, and SendOnBehalf permissions assigned to a specific user.
- Provides the option to remove permissions selectively or all at once.
- Supports Exchange Online in Office 365 environments.
- Offers flexible UPN/email input validation and error handling.

## Disclaimer
Please note that while I have taken care to ensure the script works correctly, I am not responsible for any damage or issues that may arise from its use. Use this script at your own risk.

## License
This project is licensed under the terms of the GNU General Public License v3.0.
