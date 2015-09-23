Add-Type -TypeDefinition @"
   public enum OfficeVersion
   {
      Office2003,
      Office2007,
      Office2010,
      Office2013
   }
"@

function New-OfficeAppLockGPO{

    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=0)]
        [string] $GpoName = $null,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [OfficeVersion[]] $OfficeVersion
    )

    Import-Module -Name grouppolicy

    $dateconv = Get-Date -Format G
    $date = (Get-date $dateconv).TofileTime()
   
    if(!($GpoName)){
    
        $GpoName = @("LockOffice2003","LockOffice2007","LockOffice2010","LockOffice2013")   
        $officeNumbers = @("11","12","14","15")
        $gpoCounter = 0

        foreach($Gpo in $GpoName){

            New-GPO -Name $Gpo

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

            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\Certificates" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CRLs" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CTLs" -Type String -Value ""
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\Certificates" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CRLs" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CTLs" -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName DefaultLevel -Type DWord -Value 262144 | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName ExecutableTypes -Type MultiString -Value "ADE\0ADP\0BAS\0BAT\0CHM\0CMD\0COM\0CPL\0CRT\0EXE\0HLP\0HTA\0INF\0INS\0ISP\0LNK\0MDB\0MDE\0MSC\0MSI\0MSP\0MST\0OCX\0PCD\0PIF\0REG\0SCR\0SHS\0URL\0VB\0WSC" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName PolicyScope -Type DWord -Value 0 | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName TransparentEnabled -Type DWord -Value 1 | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName Description -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName LastModified -Type QWord -Value $date | Out-Null
            Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null

            foreach($app in $appStrings)
            {
                $guid = ([system.guid]::NewGuid())
                $guidString = "{$($guid.ToString())}"

                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value $app | Out-Null   
            }

            foreach($loc in $appLocations)
            {
                $guid = ([system.guid]::NewGuid())
                $guidString = "{$($guid.ToString())}"
    
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value "{$($loc.ToString())}" | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
                Set-GPRegistryValue -Name $Gpo -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null
            }

            Write-Host "The Group Policy $Gpo has been created"

            $gpoCounter = $gpoCounter + 1
        }
    }
    
    else{
    
        New-GPO -Name $GpoName

        $appLocations = @("%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRoot%",
                          "%HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir%")

        foreach($loc in $appLocations)
        {
            $guid = ([system.guid]::NewGuid())
            $guidString = "{$($guid.ToString())}"
    
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value "{$($loc.ToString())}" | Out-Null
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
            Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\262144\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null       
        }

        if($OfficeVersion -eq "Office2003"){

            $officeNumber = '11'
            SetGpoPolValues
        }

        if($OfficeVersion -eq "Office2007"){

            $officeNumber = '12'
            SetGpoPolValues
        }

        if($OfficeVersion -eq "Office2010"){

            $officeNumber = '14'
            SetGpoPolValues
        }
                    
        if($OfficeVersion -eq "Office2013"){

            $officeNumber = '15'
            SetGpoPolValues
        }             
    }

    $results = new-object PSObject[] 0;
    $Result = New-Object –TypeName PSObject
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "GpoName" -Value $GpoName
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "WmiFilterName" -Value $WmiFilterName
    $Result
}
    
function SetGpoPolValues{
               
        $appStrings = @("C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\WINWORD.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\WINWORD.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\OUTLOOK.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\OUTLOOK.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\ONENOTE.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\ONENOTE.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\POWERPNT.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\POWERPNT.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\MSACCESS.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\MSACCESS.EXE",
                        "C:\Program Files (x86)\Microsoft Office\Office$($officeNumber)\EXCEL.EXE",
                        "C:\Program Files\Microsoft Office\Office$($officeNumber)\EXCEL.EXE")

        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\Certificates" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CRLs" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\Disallowed\CTLs" -Type String -Value ""
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\Certificates" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CRLs" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\SystemCertificates\TrustedPublisher\CTLs" -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName DefaultLevel -Type DWord -Value 262144 | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName ExecutableTypes -Type MultiString -Value "ADE\0ADP\0BAS\0BAT\0CHM\0CMD\0COM\0CPL\0CRT\0EXE\0HLP\0HTA\0INF\0INS\0ISP\0LNK\0MDB\0MDE\0MSC\0MSI\0MSP\0MST\0OCX\0PCD\0PIF\0REG\0SCR\0SHS\0URL\0VB\0WSC" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName PolicyScope -Type DWord -Value 0 | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers" -ValueName TransparentEnabled -Type DWord -Value 1 | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName Description -Type String -Value "" | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName LastModified -Type QWord -Value $date | Out-Null
        Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null

        foreach($app in $appStrings)
        {
           $guid = ([system.guid]::NewGuid())
           $guidString = "{$($guid.ToString())}"

           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName LastModified -Type QWord -Value $date | Out-Null
           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName Description -Type String -Value "" | Out-Null
           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName SaferFlags -Type DWord -Value 0 | Out-Null
           Set-GPRegistryValue -Name $GpoName -Key "HKCU\Software\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\$guidString" -ValueName ItemData -Type ExpandString -Value $app | Out-Null   
        }      
} 