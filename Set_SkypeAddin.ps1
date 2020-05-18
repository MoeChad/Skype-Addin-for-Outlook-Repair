#Requires -module ActiveDirectory

function Set-SkypeAddin {
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

    $ResiliencyPath = 'HKU:\{0}\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisturbAddinList' -f $SID
    $AddinPath = 'HKU:\{0}\Software\Microsoft\Office\16.0\Outlook\Addins\UCAddin.LyncAddin.1' -f $SID

    if ((Test-Path -Path $ResiliencyPath) -eq $false)  {
        Write-Verbose -Message ('The user path for DoNotDisturbAddinList does not exist for {0}; creating that path' -f $UserName)
        New-Item -Path $ResiliencyPath -Force > $Null
    }

    Write-Verbose -Message ('Setting properties for UCAddin.Lync.1 and LoadBehavior')
    New-ItemProperty -Path $ResiliencyPath -Name UCAddin.Lync.1 -Value 1 -PropertyType DWord -Force > $Null
    New-ItemProperty -Path $AddinPath -Name LoadBehavior -Value 3 -PropertyType DWord -Force > $Null

    Remove-PSDrive -Name HKU 
}
