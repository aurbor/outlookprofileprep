# Outlook Profile Prep Script
Powershell script to create new Default profiles for Microsoft 365 Go-Live Cutover days

# Description
I've written a Powershell script that will do the following:

- Check the HKEY_USERS hive for any Windows accounts on the PC that have an Outlook profile
- For any PC Users who have an Outlook profile, the script will create a new Outlook profile called 'O365-Username' and it will set the new profile as default.
- If the Windows Profile already has an Outlook profile called 'O365-Username' then it will remove it and create a new one.
- It then checks to see if the 'ExcludeExplicitO365Endpoint' registry key exists for any of the user accounts on the machine, and if it does, it removes it.

This script should mean that if it's run with administrative permissions on any machine that is online, it will create a new Default Outlook profile for any any active Outlook user on a PC.

# Considerations
*NOTE: THE PROFILE STILL NEEDS TO BE SETUP FOR THE USER! The script creates a blank profile, it will still be up to us (or the end user) to login with their new Office 365 username and password to complete the profile creation.*

See attached a screenshot of the script result if it's run locally on the machine. The script also logs to c:\temp. A sample log file is attached.

I've tested the script on a VM with my own Outlook profile, and it seems to work well, the main thing to think about though is that we only want to deploy the script to machines that we definitely want to kill Outlook profiles on. The old profile will still be retained in case we need to get anything from it, but it should prevent users from opening Outlook and trying to use it whilst still on Intermedia come go-live day. It's also important to note that if the script is run AFTER we re-configure a machine for the new Outlook profiles, we're effectively wiping that out by running this script, so it should really only be run once on machines that are online.
