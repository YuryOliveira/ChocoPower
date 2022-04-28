<#
.EXAMPLE
Iniciar instalação pelo teclado Win+R
PowerShell -Command "& {Start-Process PowerShell -Verb RunAs -ArgumentList {invoke-expression ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/setup.ps1'))}}"
.EXAMPLE
Iniciar instalação pelo terminal PS ou por script .ps1
Start-Process PowerShell -Verb RunAs -ArgumentList {invoke-expression ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/setup.ps1'))}}
#>
Set-ExecutionPolicy Bypass -Scope Process -Force

if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        Start-Process PowerShell -Verb RunAs -ArgumentList "invoke-expression ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/setup.ps1'))";
        Exit;
    }
}

[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

$temp = (New-item "C:\$(New-Guid)" -ItemType Directory -Force -Confirm:$false).FullName

invoke-expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

#-Headers @{"Cache-Control"="no-cache"}
$applist = (Invoke-WebRequest "https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/applist.json" -UseBasicParsing).Content | ConvertFrom-Json

function Start_Install 
{
    $applist.chocolatey | & { process {
    
        if($_.Enable -eq $True)
        {
            if(!([string]::IsNullOrEmpty($_.Args))) 
            {
                choco install $_.Name --params $_.Args --y --force
            }
            else
            {
                choco install $_.Name --y --force
            }
        }
    }}
    
    $applist.custom | & { process {
    
        if($_.Enable -eq $True)
        {
            if(!([string]::IsNullOrEmpty($_.Args))) 
            {
                Custom_Install -name $_.Name -arg $_.Args
            }
            else
            {
                Custom_Install $_.Name
            }
        }
    }}
    
    $("C:\ProgramData\chocolatey",$temp,$env:TMP) | & { process {remove-item $_ -Recurse -Force -Confirm:$false -ea 0}}
}

function Custom_Install($name,$arg)
{
    switch ($name)
    {
        "office" 
        {
            #Install Office 2021 LTSC
            $file = "$env:TMP\officedeploymenttool.exe"
            $Uri  = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
            $url  = ((Invoke-WebRequest $Uri -UseBasicParsing ).Links | Where-Object {$_.outerHTML -match "click here to download manually"}).href
            (New-Object net.webclient).Downloadfile($url, $file)
            Start-Process "$env:TMP\officedeploymenttool.exe" -ArgumentList "/quiet /extract:`"$env:TMP\officedeploymenttool`"" -NoNewWindow -Wait
            Start-Process "$env:TMP\officedeploymenttool\Setup.exe" -ArgumentList "/configure https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/deployment_office2021.xml" -NoNewWindow -Wait
        }
        "actwinoffice"
        {
            #Active Windows and Office
            $uri  = "https://api.github.com/repos/massgravel/Microsoft-Activation-Scripts/releases"
            $zip  = "MAS_*.7z"
            $url  = ((Invoke-RestMethod -Method GET -Uri $uri)[0].assets | Where-Object name -like $zip ).browser_download_url
            $file = ("$url" -split '/')[-1]
            $pass = "1234"
            try { 
                    (New-Object net.webclient).Downloadfile($url, "$temp\$file")
            }
            catch {
                    (New-Object net.webclient).Downloadfile($url, "$temp\$file")
            }
            if(Test-Path "$temp\$file") {
                Start-Process "$env:ProgramData\chocolatey\bin\7z.exe" -ArgumentList "x `"$temp\$file`" -o`"$temp`" -p`"$pass`"" -NoNewWindow -Wait
                Start-Process "$temp\MAS_*\Separate-Files-Version\Activators\HWID-KMS38_Activation\HWID_Activation.cmd" -ArgumentList "/a" -NoNewWindow -Wait #Activate Windows Digital License
                Start-Process "$temp\MAS_*\Separate-Files-Version\Activators\Online_KMS_Activation\Activate.cmd" -ArgumentList "/o" -NoNewWindow -Wait #Activate Office 180 days
                Start-Process "$temp\MAS_*\Separate-Files-Version\Activators\Online_KMS_Activation\Activate.cmd" -ArgumentList "/rt" -NoNewWindow -Wait #Create Renewal Task for Office
            }
        }
        Default {}
    }
}

Start_Install