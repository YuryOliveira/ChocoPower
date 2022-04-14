Set-ExecutionPolicy Bypass -Scope Process -Force

[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

invoke-expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

$applist = $Null
$applist = (Invoke-WebRequest "https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/applist.json" -UseBasicParsing -Headers @{"Cache-Control"="no-cache"}).Content | ConvertFrom-Json

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

    Start-Process "$env:TMP\officedeploymenttool\Setup.exe" -ArgumentList "/configure https://raw.githubusercontent.com/YuryOliveira/ChocoPower/main/deployment_office2021.xml" -wait
}

#Active Windows and Office
if($applist.actwinoffice -eq $True)
{
    $uri  = "https://api.github.com/repos/massgravel/Microsoft-Activation-Scripts/releases"
    $zip  = "MAS_*.7z"
    $url  = ((Invoke-RestMethod -Method GET -Uri $uri)[0].assets | Where-Object name -like $zip ).browser_download_url
    $file = ("$url" -split '/')[-1]
    $pass = "1234"
    
    try 
    {
        (New-Object net.webclient).Downloadfile($url, "$env:TMP\$file")
    }
    catch 
    {
        (New-Object net.webclient).Downloadfile($url, "$env:TMP\$file")
    }

    if(Test-Path "$env:TMP\$file")
    {
        Start-Process "$env:ProgramData\chocolatey\bin\7z.exe" -ArgumentList "x `"C:\Windows\Temp\$file`" -o`"$env:TMP`" -p`"$pass`"" -NoNewWindow -Wait
        Start-Process "C:\Windows\Temp\MAS_*\Separate-Files-Version\Activators\HWID-KMS38_Activation\HWID_Activation.cmd" -ArgumentList "/a" -NoNewWindow -Wait #Activate Windows Digital License
        Start-Process "C:\Windows\Temp\MAS_*\Separate-Files-Version\Activators\Online_KMS_Activation\Activate.cmd" -ArgumentList "/o" -NoNewWindow -Wait #Activate Office 180 days
        Start-Process "C:\Windows\Temp\MAS_*\Separate-Files-Version\Activators\Online_KMS_Activation\Activate.cmd" -ArgumentList "/rt" -NoNewWindow -Wait #Create Renewal Task for Office
    }
}

remove-item C:\ProgramData\chocolatey -Recurse -Force -Confirm:$false -ea 0
remove-item "C:\Windows\Temp\MAS_*" -Recurse -Force -Confirm:$false -ea 0
remove-item $env:TMP -Recurse -Force -Confirm:$false -ea 0
