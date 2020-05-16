#Prompt for username
$ID = Read-Host "Please enter username"

#Pulls domain SID from Active Directory
$Sid = (Get-ADUser -identity $ID).SID.Value

#Creates PS Drive to access HKEY USER Registry
New-PSDrive HKU Registry HKEY_USERS

#Return True/False boolean value for specified registry key
$donotdisablekey = Test-Path HKU:\$sid\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisturbAddinList

#Creates new registry key if it doesn't not already exist to repair skype add-in issue
if($donotdisablekey -eq $false) {
    New-Item HKU:\$sid\Software\Microsoft\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisturbAddinList
    New-Item HKU:\$sid\Software\Microsoft\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisturbAddinList -Name "UCaddin.Lync.1" -PropertyType DWord -value "00000001" -Force > $Null
    New-Item HKU:\$sid\Software\Microsoft\Software\Microsoft\Office\16.0\Outlook\Addins\UCAddin.1 -Name "LoadBehavior" -Propertytype Dword "00000003" -Force > $null
    }
else {
    New-Item HKU:\$sid\Software\Microsoft\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisturbAddinList -Name "UCaddin.Lync.1" -PropertyType DWord -value "00000001" -Force > $Null
    New-Item HKU:\$sid\Software\Microsoft\Software\Microsoft\Office\16.0\Outlook\Addins\UCAddin.1 -Name "LoadBehavior" -Propertytype Dword "00000003" -Force > $null
    }

#Removes HKEY USER Registry
Remove-PSDrive HKU