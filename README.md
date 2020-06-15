# Skype Add-in for Outlook Repair
This Function is designed to enable the Skype Addin in Outlook when the COM Add-in menu is not functioning correctly

### Requirements
- Active Directory module

### Functionality
1. Prompts for username
2. username is used with  NTAccount class to pull SID
3. SID is used to create the registry path to HKEY_CURRENT_USER
4. Creates the "DoNotDisableAddinList" registry key if one does not exist
5. Adds the "LoadBehavior" and "UCAddin.Lync.1" registry values to the appropriate registry keys


### Example
``` Set-SkypeAddin -Username ``` 
- you can add ``` -Verbose ``` if you would like output of what the function has done.


### Notes
- This function does have error handling with Try/Catch for the username input
