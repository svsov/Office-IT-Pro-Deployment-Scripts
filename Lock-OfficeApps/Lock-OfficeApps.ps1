Function Lock-OfficeApps{

    Import-Module -Name grouppolicy

    $GpoName = @("LockOffice2007","LockOffice2010","LockOffice2013")
    $officeNumbers = @("12","14","15")
    $gpoCounter = 0

    foreach($Gpo in $GpoName){

        New-GPO -Name $Gpo

        $dateconv = Get-Date -Format G
        $date = (Get-date $dateconv).TofileTime()

        $codeId = "HKCU:\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths"
        $appStrings = @("C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\WINWORD.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\WINWORD.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\OUTLOOK.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\OUTLOOK.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\ONENOTE.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\ONENOTE.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\POWERPNT.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\POWERPNT.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\MSACCESS.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\MSACCESS.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumbers[$gpoCounter])\EXCEL.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumbers[$gpoCounter])\EXCEL.EXE")

        $appLocations = @("%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRoot%",
                          "%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir%")

        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\Certificates" -Type String -Value ""
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CRLs" -Type String -Value ""
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CTLs" -Type String -Value ""
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\Certificates" -Type String -Value ""
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CRLs" -Type String -Value ""
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CTLs" -Type String -Value ""
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName DefaultLevel -Type DWord -Value 262144
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName ExecutableTypes -Type MultiString -Value "ADE\0ADP\0BAS\0BAT\0CHM\0CMD\0COM\0CPL\0CRT\0EXE\0HLP\0HTA\0INF\0INS\0ISP\0LNK\0MDB\0MDE\0MSC\0MSI\0MSP\0MST\0OCX\0PCD\0PIF\0REG\0SCR\0SHS\0URL\0VB\0WSC"
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName PolicyScope -Type DWord -Value 0
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName TransparentEnabled -Type DWord -Value 1
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName Description -Type String -Value ""
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName LastModified -Type QWord -Value $date
        Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName SaferFlags -Type DWord -Value 0

        foreach($app in $appStrings)
        {
            $guid = ([system.guid]::NewGuid())
            $guidString = "{$($guid.ToString())}"

            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName Description -Type String -Value ""
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value $app   
        }

        foreach($loc in $appLocations)
        {
            $guid = ([system.guid]::NewGuid())
            $guidString = "{$($guid.ToString())}"
    
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName Description -Type String -Value "" 
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value "{$($loc.ToString())}"
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0
        }
        $gpoCounter = $gpoCounter + 1
    }
}