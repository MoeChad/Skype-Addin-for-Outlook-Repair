function Set-DoNotDisableKey {
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [string]$UserName
    )

    try {
        $SID = Get-ADUser -Identity $UserName -Properties SID | Select-Object -ExpandProperty SID
        Write-Verbose -Message ('Located user {0}' -f $UserName) 
    }
    catch {
        Throw ('Unable to locate user {0}' -f $UserName)
    }

    Write-Verbose -Message 'Establishing a PSDrive to HKU'
    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS > $null 

    $UserPath = 'HKU:\{0}\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisturbAddinList' -f $SID
    if (Test-Path -Path $UserPath -eq $false)  {
        Write-Verbose -Message ('The user path for DoNotDisturbAddinList does not exist for {0}; creating that path' -f $UserName)
        New-Item -Path $UserPath -Force > $Null
    }

    Write-Verbose -Message ('Setting properties for UCAddin.Lync.1 and LoadBehavior')
    New-ItemProperty -Path $UserPath -Name UCAddin.Lync.1 -Value 1 -PropertyType DWord -Force > $Null
    New-ItemProperty -Path $UserPath -Name LoadBehavior -Value 3 -PropertyType DWord -Force > $Null

    Remove-PSDrive -Name HKU 
}