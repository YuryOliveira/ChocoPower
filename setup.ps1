Set-ExecutionPolicy Bypass -Scope Process -Force

[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

invoke-expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

$applist = (Invoke-WebRequest "https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/applist.json" -UseBasicParsing).Content | ConvertFrom-Json

$applist.chocolatey | ForEach-Object {
    
    if(!([string]::IsNullOrEmpty($_.Args))) 
    {
        choco install $_.Name --params $_.Args --y --force
    }
    else
    {
        choco install $_.Name --y --force
    }
}

#Install Office 2021 LTSC
if($applist.office -eq $True)
{
    $file = "$env:TMP\officedeploymenttool.exe"
    $Uri  = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
    $url  = ((Invoke-WebRequest $Uri -UseBasicParsing ).Links | Where-Object {$_.outerHTML -match "click here to download manually"}).href
    
    (New-Object net.webclient).Downloadfile($url, $file)

    Start-Process "$env:TMP\officedeploymenttool.exe" -ArgumentList "/quiet /extract:`"$env:TMP\officedeploymenttool`"" -Wait

    Start-Process "$env:TMP\officedeploymenttool\Setup.exe" -ArgumentList "/configure https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/deployment_office2021.xml"
}

remove-item C:\ProgramData\chocolatey -Recurse -Force
