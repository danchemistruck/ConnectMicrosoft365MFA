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
Enable-AzureRmAlias -scope localmachine
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
    if(get-module az){Enable-AzureRmAlias -scope localmachine}
    Get-AzureRmSubscription|Out-GridView -PassThru|Select-AzureRmSubscription
}

function EXO {
    BEGIN {
        $credential = Get-Password
        import-module msonline
        Start-Service WinRM
    }
    PROCESS {
        
        #Import Exchange MFA Module
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
        $ExoModulePath = [System.IO.Path]::Combine($CreateEXOPSSession, $ExoPowershellModule);
        import-module $ExoModulePath
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        . "$CreateEXOPSSession\CreateExoPSSession.ps1"

                #Connect to Exchange Online
                Write-Progress -Activity "Connecting to Exchange Online"
                try {
                    Connect-EXOPSSession -UserPrincipalName $credential.userprincipalname
                }
                Catch {
                    Write-Host 'Could not connect to Exchange online without MFA.' -foregroundcolor Magenta
                    Connect-EXOPSSession -UserPrincipalName $credential.userprincipalname
                }
    <#
        Connect-MsolService -credential $credential.creds
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        . "$CreateEXOPSSession\CreateExoPSSession.ps1"
        Connect-EXOPSSession -UserPrincipalName $credential.userprincipalname
     #>                
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
    import-module microsoftteams
    Connect-MicrosoftTeams
}

function azure {
    BEGIN {}
    PROCESS {
        import-module Az
        Enable-AzureRMAlias
        Login-AzureRmAccount
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
        $ErrorActionPreference= 'silentlycontinue'       
    }
    PROCESS {    
        #Connect to Azure AD
        Write-Progress -Activity "Connecting to Azure AD"
        try {
            Connect-AzureAD -Credential $credential.Creds
        }
        catch {
            Write-Host 'Could not connect to Azure AD without MFA.' -foregroundcolor Magenta
            Connect-AzureAD
        }
        
        #Connect to Teams
        Write-Progress -Activity "Connecting to Teams"
        try {
            Connect-MicrosoftTeams -Credential $credential.creds
        }
        catch {
            Write-Host 'Could not connect to Teams without MFA.' -foregroundcolor Magenta
            Connect-MicrosoftTeams
        }    
        #Connect to MSOL
        Write-Progress -Activity "Connecting to MSOL"
        try{
            Connect-MsolService -Credential $credential.creds
        }
        catch {
            Write-Host 'Could not connect to MSOnline without MFA.' -foregroundcolor Magenta
            Connect-MsolService -UserPrincipalName $credential.userprincipalname
        }

        $domainHost = $($(get-msoldomain |? {$_.name -like "*.onmicrosoft.com"})[0].name -split "\.")[0]
        
        #Connect to SPO
        Write-Progress -Activity "Connecting to SPO"
        Connect-SPOService -Url https://$domainhost-admin.sharepoint.com -credential $credential.creds
        
        #Connect to Skype
        Write-Progress -Activity "Connecting to Skype"
        try {
            $SboSession = New-CsOnlineSession -credential $credential.creds
            Import-PSSession $SboSession    
        }
        Catch {
            Write-Host 'Could not connect to Skype for Business Online without MFA.' -foregroundcolor Magenta
            $SboSession = New-CsOnlineSession;Import-PSSession $SboSession
        }

        #Import Exchange MFA Module
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll";
        $ExoModulePath = [System.IO.Path]::Combine($CreateEXOPSSession, $ExoPowershellModule);
        import-module $ExoModulePath
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        . "$CreateEXOPSSession\CreateExoPSSession.ps1"
        
        #Connect to Security and compliance center
        Write-Progress -Activity "Connecting to Security and Compliance Center"
        try {
            Connect-IPPSSession -UserPrincipalName $credential.userprincipalname
        }
        Catch {
            Write-Host 'Could not connect to Security and compliance center.' -foregroundcolor Magenta
            Connect-IPPSSession -UserPrincipalName $credential.userprincipalname
        }

        #Connect to Exchange Online
        Write-Progress -Activity "Connecting to Exchange Online"
        try {
            Connect-EXOPSSession -UserPrincipalName $credential.userprincipalname
        }
        Catch {
            Write-Host 'Could not connect to Exchange online without MFA.' -foregroundcolor Magenta
            Connect-EXOPSSession -UserPrincipalName $credential.userprincipalname
        }
        #Connect to Power BI
        Write-Progress -Activity "Connecting to Power BI"
        try {
            Connect-PowerBIServiceAccount -Credential $credential.Creds
        }
        Catch {
            Write-Host 'Could not connect to Power BI without MFA.' -foregroundcolor Magenta
            Connect-PowerBIServiceAccount
        }
        #Connect to Azure
        Write-Progress -Activity "Connecting to Azure"
        try{
            Connect-AzAccount -credential $credential.creds
        }
        Catch {
            Write-Host 'Could not connect to Azure without MFA.' -foregroundcolor Magenta
            Connect-AzAccount -UserPrincipalName $credential.userprincipalname
        }
        Select-AzureSub
    }
    END {}
}

function NOM365 {Get-PSSession | remove-pssession}
Set-Alias -Name NOEXO -Value NOM365

function install-m365 {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

    Install-Module -Name Az -AllowClobber -Force
    Enable-AzureRMAlias
    
    #Install all Microsoft 365 modules.
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
    Install-Module -Name Microsoft.RDInfra.RDPowerShell -force -Confirm:$false
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
