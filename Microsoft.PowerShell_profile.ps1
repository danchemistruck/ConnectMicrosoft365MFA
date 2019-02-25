<#
.SYNOPSIS
    Windows PowerShell Profile
.DESCRIPTION
    Populates your Windows PowerShell profile with functions to connect to Microsoft online services.
.EXAMPLE
    PS C:\> install-m365
    Installs all the Microsoft PowerShell modules for Microsoft 365 services.
    Note: Exchange MFA modules are in the dependencies folder of the repo and must be copied into the following folder:  %userprofile%\Documents\WindowsPowerShell
    https://github.com/danchemistruck/ConnectMicrosoft365MFA/tree/master/Dependencies
.EXAMPLE
    PS C:\> m365
    Prompts for username and password, then connects to all Microsoft 365 services.
.NOTES
    Copyright (c) Dan Chemistruck 2019. All rights reserved.

    MIT License

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the ""Software""), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>
function get-password {
    [String][ValidateScript( {$_ -match '^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$'})] $UPN = Read-Host -Prompt 'Userprincipalname'
    [SecureString] $password = Read-Host -Prompt 'Password' -AsSecureString
    $creds = New-Object System.Management.Automation.PSCredential($UPN, $password)

    $Credentials = new-object PSObject
    $Credentials | add-member -membertype NoteProperty -name "Userprincipalname" -value $UPN
    $Credentials | add-member -membertype NoteProperty -name "Password" -Value $password
    $Credentials | add-member -membertype NoteProperty -name "Creds" -Value $creds

    return $Credentials
}
function select-azuresub {
    Get-AzureRmSubscription|Out-GridView -PassThru|Select-AzureRmSubscription
}

function EXO {
    BEGIN {
        $credential = Get-Password
        import-module msonline
    }
    PROCESS {
        Connect-MsolService -credential $credential.creds
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        . "$CreateEXOPSSession\CreateExoPSSession.ps1"
        Connect-EXOPSSession -UserPrincipalName $credential.userprincipalname
    }
    END {}
}
function MSO { 
    $credential = Get-Password
    import-module msonline
    import-module AzureAD
    Connect-MsolService -credential $credential.creds
    Connect-AzureAD -Credential $credential.Creds
}
function SBO {
    $credential = Get-Password
    import-module skypeonlineconnector
    import-module AzureAD
    import-module msonline
    $session = New-CsOnlineSession -credential $credential.creds
    Import-PSSession $session
    Connect-MsolService -Credential $credential.creds
    Connect-AzureAD -Credential $credential.Creds
}
function teams {
    $credential = get-password
    import-module microsoftteams
    Connect-MicrosoftTeams -credential $credential.creds
}

function azure {
    BEGIN {}
    PROCESS {
        $credential = get-password
        import-module Az
        Enable-AzureRMAlias
        Login-AzureRmAccount -credential $credential.creds
        select-azuresub
    }
    END {}
}
function M365 {
    BEGIN {
        $credential = get-password
        $global:UserPrincipalName = $Credential.userprincipalname;
        $global:Credential = $Credential.creds;

        Start-Service WinRM
        Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1"
        import-module Az
        Enable-AzureRMAlias
        import-module AADRM
        import-module AzureAD
        Get-InstalledModule -Name MicrosoftPowerBI*| % {import-module $_.name}
        import-module microsoftteams
        import-module msonline
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking        
    }
    PROCESS {    
        #Connect to Azure AD
        Write-Host "Connecting to Azure AD"
        Connect-AzureAD -Credential $credential.Creds
        
        #Connect to Teams
        Write-Host "Connecting to Teams"
        Connect-MicrosoftTeams -Credential $credential.creds
        
        #Connect to MSOL
        Write-Host "Connecting to MSOL"
        Connect-MsolService -Credential $credential.creds
        $domainHost = $($(get-msoldomain |? {$_.name -like "*.onmicrosoft.com"})[0].name -split "\.")[0]
        
        #Connect to SPO
        Write-Host "Connecting to SPO"
        Connect-SPOService -Url https://$domainhost-admin.sharepoint.com -credential $credential.creds
        
        #Connect to Skype
        Write-Host "Connecting to Skype"
        try {
            $SboSession = New-CsOnlineSession -credential $credential.creds
            Import-PSSession $SboSession    
        }
        Catch {
            Write-host 'Could not connect to Skype for Business Online. Try running:'
            write-host '$SboSession = New-CsOnlineSession;Import-PSSession $SboSession'
        }

        #Import Exchange MFA Module
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
        $ExoModulePath = [System.IO.Path]::Combine($CreateEXOPSSession, $ExoPowershellModule);
        import-module $ExoModulePath
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        . "$CreateEXOPSSession\CreateExoPSSession.ps1"
        
        #Connect to Security and compliance center
        Write-Host "Connecting to Compliance Center"
        Connect-IPPSSession -UserPrincipalName $credential.userprincipalname
        
        #Connect to Exchange Online
        Write-Host "Connecting to EXO"
        Connect-EXOPSSession -UserPrincipalName $credential.userprincipalname

        #Connect to Power BI
        Write-Host "Connecting to Power BI"
        Connect-PowerBIServiceAccount -Credential $credential.Creds

        #Connect to Azure
        Write-Host "Connecting to Azure"
        Login-AzureRmAccount -credential $credential.creds
        Select-AzureSub
    }
    END {}
}

function NOEXO {Get-PSSession | remove-pssession}

function install-m365 {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

    Install-Module -Name Az -AllowClobber -Force
    Enable-AzureRMAlias
    
    Install-Module AADRM -force -Confirm:$false
    Install-Module AzureAD -force -Confirm:$false
    Install-Module MicrosoftTeams -force -Confirm:$false
    Install-Module msonline -force -Confirm:$false
    Install-Module -Name MicrosoftPowerBIMgmt -force -Confirm:$false
    Install-Module -Name MicrosoftPowerBIMgmt.Data -force -Confirm:$false
    Install-Module -Name MicrosoftPowerBIMgmt.Profile -force -Confirm:$false
    Install-Module -Name MicrosoftPowerBIMgmt.Reports -force -Confirm:$false
    Install-Module -Name MicrosoftPowerBIMgmt.Workspaces -force -Confirm:$false
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force -Confirm:$false
    #Install Skype Online Module
    Try{
        $FileName = "SBOModule.exe"
        $TempFolder = 'C:\Temp'
        $MMAFile = $TempFolder + "\" + $FileName    
        $URL = "https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowerShell.Exe"
        
        # Check if folder exists, if not, create it
        if (!(Test-Path $TempFolder)){
            New-Item $TempFolder -type Directory | Out-Null
        }
        # Change the location to the specified folder
        Set-Location $TempFolder
        
        # Check if Microsoft Monitoring Agent file exists, if not, download it
        if (!(Test-Path $FileName)){
            Invoke-WebRequest -Uri $URL -OutFile $MMAFile | Out-Null
        }
            
        # Install the Skype Module
        Start-Process $FileName /quiet -ErrorAction Stop -Wait | Out-Null
        
    }
    catch {
        $server = $_.Exception.Message
        $server | export-csv C:\temp\ErrorLog-SBOModule.csv -append -notypeinfo
    }
}

function update-m365 {
    Update-Module -Name Az -AllowClobber
    Update-Module -Name AADRM -Confirm:$false
    Update-Module -Name AzureAD -Confirm:$false
    Update-Module -Name MicrosoftTeams -Confirm:$false
    Update-Module -Name msonline -Confirm:$false
    Update-Module -Name MicrosoftPowerBIMgmt -Confirm:$false
    Update-Module -Name MicrosoftPowerBIMgmt.Data -Confirm:$false
    Update-Module -Name MicrosoftPowerBIMgmt.Profile -Confirm:$false
    Update-Module -Name MicrosoftPowerBIMgmt.Reports -Confirm:$false
    Update-Module -Name MicrosoftPowerBIMgmt.Workspaces -Confirm:$false
    Update-Module -Name Microsoft.Online.SharePoint.PowerShell -Confirm:$false
}

# Chocolatey profile
$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
if (Test-Path($ChocolateyProfile)) {
    Import-Module "$ChocolateyProfile"
}
