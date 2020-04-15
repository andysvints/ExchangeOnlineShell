<#
.Synopsis Override Get-PSImplicitRemotingSession function for reconnection
#>
function global:UpdateImplicitRemotingHandler()
{
    $modules = Get-Module tmp_*

    foreach ($module in $modules)
    {
        [bool]$moduleProcessed = $false
        [string] $moduleUrl = $module.Description
        [int] $queryStringIndex = $moduleUrl.IndexOf("?")
        $ExchangeOnlineShellModule=Get-Module ExchangeOnlineShell
        $ModuleBase=$ExchangeOnlineShellModule.ModuleBase
        Import-Module "$ModuleBase\Microsoft.Exchange.Management.ExoPowershellModule.dll" -Global
        Import-Module "$ModuleBase\Microsoft.IdentityModel.Clients.ActiveDirectory.dll" -Global
        Import-Module "$ModuleBase\Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll" -Global
        
        if ($queryStringIndex -gt 0)
        {
            $moduleUrl = $moduleUrl.SubString(0,$queryStringIndex)
        }

        if ($moduleUrl.EndsWith("/PowerShell-LiveId/", [StringComparison]::OrdinalIgnoreCase) -or $moduleUrl.EndsWith("/PowerShell", [StringComparison]::OrdinalIgnoreCase))
        {
            & $module { ${function:Get-PSImplicitRemotingSession} = `
            {
                param(
                    [Parameter(Mandatory = $true, Position = 0)]
                    [string]
                    $commandName
                )

                if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened'))
                {
                    Set-PSImplicitRemotingSession `
                        (& $script:GetPSSession `
                            -InstanceId $script:PSSession.InstanceId.Guid `
                            -ErrorAction SilentlyContinue )
                }
                if (($script:PSSession -ne $null) -and ($script:PSSession.Runspace.RunspaceStateInfo.State -eq 'Disconnected'))
                {
                    # If we are handed a disconnected session, try re-connecting it before creating a new session.
                    Set-PSImplicitRemotingSession `
                        (& $script:ConnectPSSession `
                            -Session $script:PSSession `
                            -ErrorAction SilentlyContinue)
                }
                if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened'))
                {
                    Write-PSImplicitRemotingMessage ('Creating a new Remote PowerShell session using MFA for implicit remoting of "{0}" command ...' -f $commandName)
                    $session = New-ExoPSSession -UserPrincipalName $global:UserPrincipalName -ConnectionUri $global:ConnectionUri -AzureADAuthorizationEndpointUri $global:AzureADAuthorizationEndpointUri -PSSessionOption $global:PSSessionOption -Credential $global:Credential

                    if ($session -ne $null)
                    {
                        Set-PSImplicitRemotingSession -CreatedByModule $true -PSSession $session
                    }

                    RemoveBrokenOrClosedPSSession
                }
                if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened'))
                {
                    throw 'No session has been associated with this implicit remoting module'
                }

                return [Management.Automation.Runspaces.PSSession]$script:PSSession
            }}
        }
    }
}

<#
.Synopsis Remove broken and closed sessions
#>
function  global:RemoveBrokenOrClosedPSSession()
{
    $psBroken = Get-PSSession | where-object {$_.State -like "*Broken*"}
    $psClosed = Get-PSSession | where-object {$_.State -like "*Closed*"}

    if ($psBroken.count -gt 0)
    {
        for ($index = 0; $index -lt $psBroken.count; $index++)
        {
            Remove-PSSession -session $psBroken[$index]
        }
    }

    if ($psClosed.count -gt 0)
    {
        for ($index = 0; $index -lt $psClosed.count; $index++)
        {
            Remove-PSSession -session $psClosed[$index]
        }
    }
}

<#
.Synopsis
   Connect to Exchange Online Powershell using Proxy settings or directly.
.DESCRIPTION
   Used to Connect to ExchangeOnlinePowershell. 
   It checks if you computer is using any proxy settings and import them from IE if needed.
.EXAMPLE
   Connect-ExchangeOnlineShell 
   Connecting to Exchange Online Shell and let the cmdlet to figure out if any Proxy Settings are in place.
.EXAMPLE
   $credObject = Get-AutomationPSCredential -Name 'MSOnline-Credentials'
   PS C:\>Connect-ExchangeOnlineShell -Credential $credObject 
   Connecting to Exchange Online Shell using credentials stored in $$credObject variable and let the cmdlet to figure out if any Proxy Settings are in place. Can be used in Azure automation runbooks.
.EXAMPLE 
   Connect-ExchangeOnlineShell -SkipProxyCheck
   Connecting to Exchange Online Shell directly
.EXAMPLE
    Connect-ExchangeOnlineShell -RenameConsoleWindow
    Connecting to Exchange Online Shell and renaming console windows to "Exchange Mangement Shell: <DefaultDomain>".
.EXAMPLE
    Connect-ExchangeOnlineShell -IdleSessionTimeout Max
    
    SessionID       : 3
    Name            : EOShell - domain.com
    ComputerName    : outlook.office365.com
    EOPrimaryDomain : domain.com
    IdleTimeout     : 43200000
    DateCreated     : 08/23/2019 10:54:37
    
    Connecting to Exchange Online Shell and setting Session Timeout to Maximum value - 12 hours.
#>
function Connect-ExchangeOnlineShell
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [Alias('ceos','Connect-EOShell')]
    Param
    (
        # Credentials Used to Connect to Exchange Online Shell
        [System.Management.Automation.PSCredential]$Credential= $null,
        
        #Used to Skip Proxy Settings Check
        [Alias('NoProxy')]
        [Switch]$SkipProxyCheck,

        #Used to rename console window to 'Exchange Online Management Shell: <default domain>'
        [Alias()]
        [switch]$RenameConsoleWindow,

        # Used to request a new URI when connecting to EXO
        [Alias('SetEmail')]
        [string]$EmailAddress,

         #Used to set custom IdleSessionTimeout in miliseconds. Default value 15 mins(900000 miliseconds), max value is 12 hours(43200000)
        [ValidatePattern({^[900000-43200000]|max*$})]
        [Alias('SessionTimeout','Timeout')]
        $IdleSessionTimeout=900000,

        #Used to connect to Government Cloud 
        [Alias('GCCHigh','Gov','GCCH')]
        [Switch]$GovernmentCloud

    )

    Begin
    {
        if ($GCCH)
        {
            $DomainSuffix="us"
        }else{
            $DomainSuffix="com"
        }

        $ConnectionUri="https://outlook.office365.$DomainSuffix/powershell-liveid/?email=$EmailAddress"
        $AzureADAuthorizationEndpointUri="https://login.windows.net/common"
        $global:ConnectionUri = $ConnectionUri;
        $global:AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri;
        $global:UserPrincipalName = $UserPrincipalName;
        
        $global:Credential = $Credential;
        $ProxyUsed=$false
        if($IdleSessionTimeout -eq "Max"){
            $IdleSessionTimeout=43200000
        }
        if(!$SkipProxyCheck){
            
            Write-Verbose "Checking Proxy Settings on computer it might take some time. Please be patient"
            Write-Verbose "Proxy Setings Check 1 of 2: Registry Check"
            Write-Verbose "Reading Proxy Settings from Registry"
            $Proxy = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
            if($Proxy.ProxyEnable){
                Write-Verbose "Proxy Settings have been detected"
                $ProxyUsed=$true
                break
            }else{
                Write-Verbose "No Proxy Settings detected"
                $ProxyUsed=$false
            }

            Write-Verbose "Proxy Setings Check 2 of 2: Transparent Proxy Check"
            Write-Verbose "If you have a transparent proxy, your computer will not provide any info about it"
            Write-Verbose "Trying to connect to outlook.office365.com via port 443 directly"
            $TCPobj=New-Object System.Net.Sockets.TCPClient
            $Connect=$TCPobj.BeginConnect("outlook.office365.com",443,$null,$null)
            $wait=$Connect.AsyncWaitHandle.WaitOne(2000,$false)
            if(!$wait){
                Write-Verbose "Connection could not be established : TimeOut has been reached"
                Write-Verbose "Most likely your computer is using Transparent Proxy"
                $ProxyUsed=$true
            }else{
                $TCPobj.EndConnect($Connect) | out-Null 
                Write-Verbose "Connection has been established successfully"
                Write-Verbose "No Transparent Proxy settings detected"
                $ProxyUsed=$false
            }
            
        }else{
            Write-Verbose "Skipping Proxy Settings Check and Connecting to outlook.office365.com directly"
        }
    
    }
    Process
    {
        if ($pscmdlet.ShouldProcess( "Exchange Online Management Shell")){
            $ExchangeOnlineSessionObjectError=$false
            if($ProxyUsed){
                Write-Verbose "Proxy Server is used: Importing Proxy Settings from IE"
                $proxySettings = New-PSSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic -IdleTimeout $IdleSessionTimeout
                $global:PSSessionOption = $proxySettings;
                Write-Verbose "Creating Session Object using $($Credential.UserName) credentials"
                
                try{
                    $PSSession=New-ExoPSSession -UserPrincipalName $($Credential.UserName) -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $proxySettings -Credential $Credential
                }catch{
                    $ExchangeOnlineSessionObjectError=$true
                    Write-Error "Catched Exception: $($_.exception.message)"
                }
            }else{
                Write-Verbose "No Proxy Detected: Connecting to Exchange Online Shell Directly"
                Write-Verbose "Creating Session Object using $($Credential.UserName) credentials"
                
                
                try{
                    $SessionSettings = New-PSSessionOption -IdleTimeout $IdleSessionTimeout
                    $PSSession=New-ExoPSSession -UserPrincipalName $($Credential.UserName) -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri  -Credential $Credential -PSSessionOption $SessionSettings
                }catch{
                    $ExchangeOnlineSessionObjectError=$true
                    Write-Error "Catched Exception: $($_.exception.message)"
                }
            }

            if($ExchangeOnlineSessionObjectError){
                Write-Error "Failed to create a Session Object"
                Write-Error "Please double check your credentials and try again"
            }else{
                Write-Verbose "Session Object has been created successfully"
                Write-Verbose "Importing Created Session"
                try{
                    Import-Module (Import-PSSession $PSSession -AllowClobber -ErrorAction SilentlyContinue) -Global -ErrorAction Stop -ErrorVariable $ImportSessionObjectError
                    Write-Verbose "Session has been imported successfully"
                    Write-Verbose "Now you are connected to ExchangeOnlieShell"
                    $Domain=Get-AcceptedDomain | Where {$_.Default -eq $true} | select -ExpandProperty DomainName -ErrorAction SilentlyContinue
                    $PSSession.Name="EOShell - $Domain"
                    $props=@{
                        SessionID=$($PSSession.Id)
                        Name="EOShell - $Domain"
                        ComputerName=$PSSession.ComputerName
                        EOPrimaryDomain=$Domain
                        DateCreated=$(get-date -Format "MM/dd/yyyy HH:mm:ss")
                        IdleTimeout=$($PSSession.IdleTimeout)
                    }
                    $EOSession=New-Object -TypeName psobject -Property $props
                    if($global:EOShellEstablishedSession){
                        $global:EOShellEstablishedSession.Add($EOSession) | Out-Null
                    }else{
                        $global:EOShellEstablishedSession=New-Object System.Collections.ArrayList
                        $global:EOShellEstablishedSession.Add($EOSession) | Out-Null
                    }
                    Write-Verbose $PSSession
                    $EOSession |Select-Object SessionID,Name,ComputerName,EOPrimaryDomain,IdleTimeout,DateCreated
                    UpdateImplicitRemotingHandler
                }catch{
                    Write-Error "Catched Exception: $($_.exception.message)"
                }
                
            }
       }
    }
    End
    {
        
        If($RenameConsoleWindow){
            Write-Verbose "Renaming Console Window to Exchange Online Management Shell: $Domain"
            $HOst.UI.RawUI.WindowTitle="Exchange Online Management Shell: $Domain"
        }
    }
}

<#
.Synopsis
   Disconnect the Exchange Online Session.
.DESCRIPTION
   Gets all the Sessions with Microsoft.Exchange Configuration established to outlook.office365.com and removes them. 
.EXAMPLE
    Disconnect-ExchangeOnlineShell
.EXAMPLE
    Disconnect-ExchangeOnlineShell

    There is 1 Exchange Online Sessions:

    SessionID       : 3
    Name            : EOShell - exelegent.com
    ComputerName    : outlook.office365.com
    EOPrimaryDomain : exelegent.com
    IdleTimeout     : 43200000
    DateCreated     : 08/23/2019 10:54:37
#>
function Disconnect-ExchangeOnlineShell
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [Alias('deos','Kill-ExchangeOnlineShellSession','Kill-EOShellSession','Disconnect-EOShell')]
    Param
    (
        #Used to Specify Domain Name for Session You want to close
        [Alias('')]
        $DomainName,
        
        #Used to Specify SessionID You want to close
        [Alias('ID')]
        $SessionID


    )

    Begin
    {
      
    }
    Process
    {
        if ($pscmdlet.ShouldProcess( "Exchange Online Powershell Sessions")){
            if($DomainName){
                try{
                    Write-Verbose "Disconnection Session with Domain Name - $DomainName"
                    $SessionToClose=$Global:EOShellEstablishedSession | Where-Object {$_.EOPrimaryDomain -eq $DomainName}
                    Get-PSSession -ID $($SessionToClose.SessionID)| Remove-PSSession
                    $Global:EOShellEstablishedSession.Remove($SessionToClose)
                }catch{
                    Write-Error "Catched Exception: $($_.exception.message)"    
                }
            }elseif($SessionID){
                try{
                    Write-Verbose "Disconnection Session with SessionID - $SessionID"
                    $SessionToClose=($Global:EOShellEstablishedSession | Where-Object {$_.SessionID -eq $SessionID})
                    Get-PSSession -ID $($SessionToClose.SessionID)| Remove-PSSession
                    $Global:EOShellEstablishedSession.Remove($SessionToClose)
                }catch{
                    Write-Error "Catched Exception: $($_.exception.message)"    
                }
            }else{
                
                 if($($Global:EOShellEstablishedSession.Count) -eq 1){
                        Write-Host "There is $($Global:EOShellEstablishedSession.Count) Exchange Online Sessions:"
                        Show-EOShellSession
                        Get-PSSession -ID $($Global:EOShellEstablishedSession.SessionID)| Remove-PSSession
                        $Global:EOShellEstablishedSession.Remove($Global:EOShellEstablishedSession[0])
                    }else{
                        Write-Host "There are $($Global:EOShellEstablishedSession.Count) Exchange Online Sessions:"
                        Show-EOShellSession
                        Write-Host ""
                        Write-Host "Please enter SessionID or DomainName to close individual session. Please enter 'All' to close all available sessions."
                        $SessionInput=Read-Host "SessionID or DomainName"
                        if($SessionInput -eq "All" -or $SessionInput -eq "ALL" -or $SessionInput -eq "all"){
                            Write-Verbose "Closing All available sessions"
                            foreach($s in $Global:EOShellEstablishedSession){
                                Remove-PSSession -Id $($s.SessionID)
                                $Global:EOShellEstablishedSession.Remove($s)
                            }
                            
                        }else{
                            if(IsInt -Text $SessionInput){
                                try{
                                    Write-Verbose "Closing Session with SessionID - $SessionInput"
                                    $SessionToClose=$(($Global:EOShellEstablishedSession | Where-Object {$_.SessionID -eq $SessionInput})) 
                                    Get-PSSession -Id $($SessionToClose.SessionID) | Remove-PSSession
                                    $Global:EOShellEstablishedSession.Remove($SessionToClose)
                                }catch{
                                    Write-Error "Catched Exception: $($_.exception.message)"
                                }
                                
                            }else{
                                
                                try{
                                    Write-Verbose "Closing Session with DomainName - $SessionInput"
                                    $SessionToClose=$(($Global:EOShellEstablishedSession | Where-Object {$_.DomainName -eq $SessionInput})) 
                                    Get-PSSession -Id $($SessionToClose.SessionID) | Remove-PSSession 
                                    $Global:EOShellEstablishedSession.Remove($SessionToClose)
                                }catch{
                                    Write-Error "Catched Exception: $($_.exception.message)"
                                }
                            
                            }
                        }
                    }
                
            }
        }
        
    }
    End
    {
        if($Host.UI.RawUI.WindowTitle -like "*Exchange Online Management Shell:*"){
            $Host.UI.RawUI.WindowTitle="Windows PowerShell"
        }
    }
}


<#
.Synopsis
   Show Established Powershell Sessions to Exchange Online.
.DESCRIPTION
   Show All Established Powershell Sessions to Exchange Online.
.EXAMPLE
   Show-EOShellSession
   
    SessionID       : 3
    Name            : EOShell - primarydomain.com
    ComputerName    : outlook.office365.com
    EOPrimaryDomain : primarydomain.com
    IdleTimeout     : 43200000
    DateCreated     : 08/23/2019 10:54:37
    
    Showing established sessions to ExchangeOnlineShell.

#>
function Show-EOShellSession
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [Alias('seos')]
    [OutputType([String])]
    Param
    (
    
    )

    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("Target", "Operation"))
        {
             foreach($s in $Global:EOShellEstablishedSession){
                    Write-Output $s | Select-Object SessionID,Name,ComputerName,EOPrimaryDomain,IdleTimeout,DateCreated
                }
        }
    }
    End
    {
    }
}


<#
.Synopsis
   Check if entered string is converatble to Integer.
.DESCRIPTION
   Check if entered string is converatble to Integer.
.EXAMPLE
   PS C:\> IsInt -Text "-1"
True
.EXAMPLE
   PS C:\> Convertable-ToInt -Text "-1qw"
False
#>
function Convertable-ToInt
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [Alias('IsInt')]
    [OutputType([String])]
    Param
    (
        # Text to try to convert to int
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("Input")] 
        $Text
      
    )

    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("$Text"))
        {
            $IsNumber=$false
            try
            {
                $IntText=[int]$Text
                $IsNumber=$true

            }
            catch
            {
                $IsNumber=$false
            }
            $IsNumber       
        }
    }
    End
    {
    }
}



