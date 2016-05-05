<#
.SYNOPSIS
    Script for updating the Office 365 SpamExperts Transport Rule.
    Schedule this script each day to ensure that the Transport Rule always has uptodate IP addresses of SpamExperts mailservers
.AUTHOR 
    Bart Tacken - Client ICT Groep
.PREREQUISITES
    Password file in XML format, follow below to create one: 
        $User = "username"
        $PassWord = "P@ssword"
        $PassWord = $PassWord | ConvertTo-SecureString -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PsCredential($user, $Password)
        $Credential | Export-Clixml $path
    PowerShell version 3
.EXAMPLE
    Script must be run from a View connection broker with following parameters:
    .\Set-SpamXpertsO365TransportRule -TransportRuleName TEST_spamexperts -CredPath c:\temp\cred4.xml
 #>
    Param(
        [Parameter(Mandatory=$True)
        ]
        [String]$TransportRuleName,
    
        [Parameter(Mandatory=$True)
        ]
        [String]$CredPath    
    )

    Function Connect-EXOnline {
        param($Credentials)
        $URL = "https://ps.outlook.com/powershell"  
        If ($Credentials -eq $Null) {
            $Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
        }
        $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
            Import-PSSession $EXOSession
    }

    # Variables
    $TransportRuleName = "TEST_spamexperts" 
    $Site = "http://noc.spamexperts.net"
    $Transcript = "C:\Windows\Temp\Set-SpamXpertsO365Connector.log"
    $Array = @()
    $credXML = Import-Clixml $CredPath
    $SiteContent = $(Invoke-WebRequest -Uri $Site).rawcontent  
    $Credential = New-Object System.Management.Automation.PsCredential($credXML.UserName, $credXML.Password) # Create PScredential Object

    Start-Transcript -Path $Transcript -Force
    Connect-EXOnline -Credentials $CredXML

    # REGex for IPv4 and IPv6 addresses
    $REGexIP = "^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$|^(([a-zA-Z]|[a-zA-Z][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z]|[A-Za-z][A-Za-z0-9\-]*[A-Za-z0-9])$|^\s*((([0-9A-Fa-f]{1,4}:){7}([0-9A-Fa-f]{1,4}|:))|(([0-9A-Fa-f]{1,4}:){6}(:[0-9A-Fa-f]{1,4}|((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){5}(((:[0-9A-Fa-f]{1,4}){1,2})|:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){4}(((:[0-9A-Fa-f]{1,4}){1,3})|((:[0-9A-Fa-f]{1,4})?:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){3}(((:[0-9A-Fa-f]{1,4}){1,4})|((:[0-9A-Fa-f]{1,4}){0,2}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){2}(((:[0-9A-Fa-f]{1,4}){1,5})|((:[0-9A-Fa-f]{1,4}){0,3}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){1}(((:[0-9A-Fa-f]{1,4}){1,6})|((:[0-9A-Fa-f]{1,4}){0,4}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(:(((:[0-9A-Fa-f]{1,4}){1,7})|((:[0-9A-Fa-f]{1,4}){0,5}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:)))(%.+)?$"
    #regex with no end of string so it captures whole line:
    $REGexFullLine =  "^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$|^(([a-zA-Z]|[a-zA-Z][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z]|[A-Za-z][A-Za-z0-9\-]*[A-Za-z0-9])$|^\s*((([0-9A-Fa-f]{1,4}:){7}([0-9A-Fa-f]{1,4}|:))|(([0-9A-Fa-f]{1,4}:){6}(:[0-9A-Fa-f]{1,4}|((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){5}(((:[0-9A-Fa-f]{1,4}){1,2})|:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){4}(((:[0-9A-Fa-f]{1,4}){1,3})|((:[0-9A-Fa-f]{1,4})?:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){3}(((:[0-9A-Fa-f]{1,4}){1,4})|((:[0-9A-Fa-f]{1,4}){0,2}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){2}(((:[0-9A-Fa-f]{1,4}){1,5})|((:[0-9A-Fa-f]{1,4}){0,3}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){1}(((:[0-9A-Fa-f]{1,4}){1,6})|((:[0-9A-Fa-f]{1,4}){0,4}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(:(((:[0-9A-Fa-f]{1,4}){1,7})|((:[0-9A-Fa-f]{1,4}){0,5}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:)))(%.+)?"

    ForEach($Row in $SiteContent.split("`n")) { # Loop through each line of Raw content     
        If ($Row -match $REGexIP) { # Add IP address to array 
            write-host $Row
            $Array = $Array + $Row        
        } # End If $REGex Match
        ElseIf ($Row -match $REGexFullLine) { # Capture IP address if first characters match with IPv4 or IPv6 address, don't touch other characters after it 
            Write-Host "Matches REGexFullLine: [$Row]"
            $Array = $Array + $Matches[0] # Add to array
        } # End if REGexFullLine        
    } # End ForEach 
    $ArrayUnique= $Array | Sort-Object -Unique

    #Update O365 Transport rule
    Get-TransportRule -Identity $TransportRuleName | Select-Object -ExpandProperty SenderIPranges # Multi valued property # Log current IP addresses
    Set-TransportRule -Identity $TransportRuleName -SenderIpRanges $ArrayUnique -verbose
    
    Stop-Transcript
