A collection of Azure Automation runbooks for Microsoft 365 and Intune management.

##Overview
This repository contains PowerShell scripts designed to be used as Azure Automation runbooks for automating various Microsoft 365 and Intune management tasks. These scripts help streamline administrative processes, maintain consistency across your environment, and reduce manual overhead.

##Repository Structure
The repository is organized into folders, with each folder containing a specific runbook solution:

Azure-Runbooks/
├── DeviceCategorySync/             # Sync device categories with user departments
├── Report-DiscoveredApps/          # Generate reports of discovered applications
├── Report-IntuneDeviceCompliance/  # Generate device compliance reports
├── Report-DevicesWithApp/          # Find devices with specific applications
├── Alert-DeviceSyncReminder/       # Send reminders for devices needing sync
├── Update-AutopilotDeviceGroupTags/ # Sync Autopilot group tags with Intune categories
├── Alert-IntuneAppleTokenMonitor/  # Monitor Apple token expirations
├── Report-UserManagers/            # Generate reports of users and their managers
├── Report-MissingSecurityUpdates/  # Report on devices missing security updates
├── Sync-IntuneDevices/             # Force sync all managed Intune devices
├── Report-DeviceSyncOverdue/       # Report on devices overdue for sync
├── Report-OneDriveSharedItems/     # Generate reports of shared items in OneDrive
├── Task-SetCompanyAttribute/       # Set company attribute for all users
├── Snipe-IT-UserSync/              # Sync Microsoft 365 users to Snipe-IT users
├── Sync-IntuneToAction1Categories/ # Sync Intune device categories to Action1 custom attributes

Each runbook folder contains:
main PowerShell script (.ps1)
A helper script for setting up permissions (Add-GraphPermissions.ps1)
Detailed documentation (README.md)

##Authentication
All runbooks in this repository are designed to use Azure Automation's System-Assigned Managed Identity for authentication, which is the recommended approach for Azure Automation. Each folder includes an Add-GraphPermissions.ps1 script that helps assign the necessary Microsoft Graph API permissions to your Automation Account's Managed Identity.

##Getting Started
Each runbook includes detailed documentation for implementation and usage. In general, to use these runbooks:

Import the script into your Azure Automation account
Enable System-Assigned Managed Identity on your Automation account
Use the included Add-GraphPermissions.ps1 script to assign necessary Graph API permissions
Configure any required parameters specific to your environment
Schedule the runbook or run it on-demand as needed

#Available Runbooks
##Reporting
Device Compliance Report: Generate comprehensive compliance reports for Intune-managed devices.
Discovered Apps Report: Create detailed reports of applications discovered on managed devices.
Devices with Specific App Report: Identify all devices with a specific application installed.
User Managers Report: Generate a report of all licensed internal users along with their manager information.
Missing Security Updates Report: Identify Windows devices missing multiple security updates with automated reporting.
Device Sync Overdue Report: Generate reports of devices that haven't synced within a specified threshold.
OneDrive Shared Items Report: Create reports of items shared externally in OneDrive for Business.
Device Management
Device Category Sync: Automatically synchronize Intune device categories based on user department information.
Autopilot Group Tag Sync: Keep Autopilot device group tags in sync with Intune device categories.
Force Device Sync: Initiate sync commands for all managed Intune devices with batching and throttling protection.
Intune to Action1 Category Sync: Sync Intune device categories to Action1 custom attributes by matching serial numbers.
Alerts and Notifications
Device Sync Reminder: Send automated email reminders to users whose devices haven't synced recently.
Apple Token Monitor: Monitor and alert on Apple Push Notification certificate and token expirations.
User Management
Company Attribute Setting: Set a consistent company attribute across all user accounts in your Microsoft 365 tenant.
Snipe-IT User Sync: Create or update Snipe-IT users from your Microsoft 365 tenant, using email as the anchor, with secure passwords for new users and optional login/invite toggles.
Third-Party Integration
Action1 Integration: Sync Intune device categories to Action1 RMM custom attributes for unified device management across platforms.

##Branch Management
This repository follows a simplified Git workflow:

The main branch contains stable, production-ready scripts
Development branches are created for new features or significant modifications
Once development work is merged into main, the development branches are typically deleted
For users who have cloned this repository, note that development branches may disappear after their work is completed
If you're working with a specific development branch, consider creating your own fork to ensure your work isn't affected when branches are deleted.

##What's New in v1.4.0
New Runbook: Sync-IntuneToAction1Categories
This release introduces a new integration with Action1 RMM, enabling automated synchronization of Intune device categories to Action1 custom attributes. Key features include:

Automatic matching of devices between Intune and Action1 using serial numbers
Syncs Intune device categories to configurable Action1 custom attributes
Supports multiple Action1 regions (North America, Europe, Australia)
Secure credential management through Azure Automation encrypted variables
WhatIf mode for testing without making changes
Comprehensive logging and statistics
Built on the PSAction1 PowerShell module
This integration helps organizations maintain consistent device categorization across both Microsoft Intune and Action1 RMM platforms, enabling better reporting, policy application, and device management workflows.

##Discussions
I've enabled GitHub Discussions for this repository to foster collaboration and support among users. This is the best place to:

Ask questions about implementing specific runbooks
Share your success stories and implementations
Suggest new runbook ideas or improvements
Discuss best practices for Azure Automation
Get help with troubleshooting
Check out the Discussions tab to join the conversation. We encourage you to use Discussions for general questions and community interaction, while Issues should be used for reporting bugs or specific problems with the scripts.

##Contributing
Feel free to use these scripts as a starting point for your own automation needs. Contributions, improvements, and suggestions are welcome!

##License
This project is licensed under the MIT License - see the LICENSE file for details.
