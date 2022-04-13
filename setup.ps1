Set-ExecutionPolicy Bypass -Scope Process -Force

[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

invoke-expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

$SoftwareList = "https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/softwares.json"

(Invoke-WebRequest $SoftwareList -UseBasicParsing).Content | ConvertFrom-Json | ForEach-Object {
    
    if(!([string]::IsNullOrEmpty($_.Args))) 
    {
        choco install $_.Name -Param $_.Args --y --force
    }
    else
    {
        choco install $_.Name
    }
}
