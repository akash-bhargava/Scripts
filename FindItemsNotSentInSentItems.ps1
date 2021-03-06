# FindItemsNotSentInSentItems.ps1
# Finds Items not sent in the sent Items folder.
# The script requires the EWS managed API, which can be downloaded here:
# https://www.microsoft.com/en-us/download/details.aspx?id=42951

#DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.

# To search multiple mailboxes with Impersonation. Add all the Primary SMTP address to the Users.csv
# Example:
#.\FindItemsNotSentInSentItems.ps1 -Username serviceaccount -Password password -Domain domain -EwsUrl "https://mail.contoso.com/ews/exchange.asmx" -IgnoreSSLCertificate -Impersonate

# Example: 
# $cred =Get-Credential
# .\FindItemsNotSentInSentItems.ps1 -Credentials $cred -EwsUrl "https://mail.contoso.com/ews/exchange.asmx" -IgnoreSSLCertificate


param (
	[Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")]
	[string]$Username,
	
	[Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")]
	[string]$Password,
	
	[Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")]
	[string]$Domain,
	
	[Parameter(Mandatory=$False,HelpMessage="Credentials")]
	[PSCredential]$Credentials,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")]
	[switch]$Impersonate,
	
	[Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]	
	[string]$EwsUrl,
	
	[Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]	
	[string]$EWSManagedApiPath = "",
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]	
	[switch]$IgnoreSSLCertificate,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]	
	[switch]$AllowInsecureRedirection,
	
	[Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	[string]$LogFile = "",
	
	[Parameter(Mandatory=$False,HelpMessage="Show Request and Response.")]	
	[switch]$RequestResponse
)


[string]$info = "White"                # Color for informational messages
[string]$warning = "Yellow"            # Color for warning messages
[string]$error = "Red"                # Color for error messages
[string]$success = "Green"                # Color for success messages
[string]$LogFile = "UnsentItems.txt"   # Path of the Log File

$verbose =$true

# Define our functions

Function Log([string]$Details, [ConsoleColor]$Colour)
{
    if ($Colour -eq $null)
    {
        $Colour = [ConsoleColor]::White
    }
	
	if ($verbose)
	{	Write-Host $Details -ForegroundColor $Colour }
		
	if ( $LogFile -eq "" ) { return	}
	(Get-Date).ToString()+" "+ $Details | Out-File $LogFile -Append
}

Function LoadEWSManagedAPI()
{
	# Find and load the managed API
	
	if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		if ( Test-Path $EWSManagedApiPath )
		{
			Add-Type -Path $EWSManagedApiPath
			return $true
		}
		Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) -ForegroundColor Yellow
	}
	
	$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	if (!$a)
	{
		$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	}
	
	if ($a)	
	{
		# Load EWS Managed API
		Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray
		Add-Type -Path $a.VersionInfo.FileName
		return $true
	}
	return $false
}

Function TrustAllCerts() {
    <#
    .SYNOPSIS
    Set certificate trust policy to trust self-signed certificates (for test servers).
    #>

    ## Code From http://poshcode.org/624
    ## Create a compilation environment
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    $Compiler=$Provider.CreateCompiler()
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters
    $Params.GenerateExecutable=$False
    $Params.GenerateInMemory=$True
    $Params.IncludeDebugInformation=$False
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

    $TASource=@'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            { 
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
                return true;
            }
        }
        }
'@ 
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
    $TAAssembly=$TAResults.CompiledAssembly

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
}

function CreateService($targetMailbox)
{
    # Creates and returns an ExchangeService object to be used to access mailboxes
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

    # Set credentials if specified, or use logged on user.
    if ($Null -ne $Credentials)
    {
        Write-Host "Applying given credentials"
        $exchangeService.Credentials = $Credentials.GetNetworkCredential()
    }
    elseif ($Username -and $Password)
    {
	    Write-Host "Applying given credentials for $Username"
	    if ($Domain)
	    {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
	    } else {
		    $exchangeService.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password)
	    }
    }
    else
    {
	    Write-Host "Using default credentials"
        $exchangeService.UseDefaultCredentials = $true
    }

    # Set EWS URL if specified, or use autodiscover if no URL specified.
    if ($EwsUrl)
    {
    	$exchangeService.URL = New-Object Uri($EwsUrl)
    }
    else
    {
    	try
    	{
		    Write-Host "Performing autodiscover for $targetMailbox"
		    if ( $AllowInsecureRedirection )
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox, {$True})
		    }
		    else
		    {
			    $exchangeService.AutodiscoverUrl($targetMailbox)
		    }
		    if ([string]::IsNullOrEmpty($exchangeService.Url))
		    {
			    Log "$targetMailbox : autodiscover failed" Red
			    return $Null
		    }
		    Write-Host "EWS Url found: $($exchangeService.Url)"
    	}
    	catch
    	{
    	}
    }
 
    if ($Impersonate)
    {
		$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $targetMailbox)
	}

	if($RequestResponse)
	{
		$exchangeService.TraceEnabled = $true
	}
	
    return $exchangeService
}

function SearchSentItems($MailboxName)
{
	#Process the mailbox
	
	Log "Processing mailbox '$($MailboxName)'"
	
	$global:service=CreateService($MailboxName)

	if($Null -eq $global:service)
	{
		Write-Host "FailedtocreateExchangeService"-ForegroundColor Red
	}
	
	#Number of Items to Get
	$pageSize =50
	$Offset = 0
	$MoreItems =$true

	$PR_MESSAGE_FLAGS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition("0x0E07", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
	
	while ($MoreItems)
	 {
	 	#Setup the View to get a limited number of Items at one time
		$itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView($pageSize,$Offset,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
		$itemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
		$itemView.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet(
							[Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
							[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass,
							[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject,
							$PR_MESSAGE_FLAGS);

							#Create the Search Filter.
		$oSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_MESSAGE_FLAGS, 1033)
		
		$oFindItems = $service.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$oSearchFilter,$itemView)

		Log "Items Found:'$($oFindItems.Items.Count)'" $info
		
	 if ($oFindItems.MoreAvailable -eq $false)
		{$MoreItems = $false}
		
	 }
	Log "Finished with Mailbox Name: '$($MailboxName)'" $info

	$global:service=$Null
}

# The following is the main script

# Check if we need to ignore any certificate errors
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!)
if ($IgnoreSSLCertificate)
{
	Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow
    TrustAllCerts
}
 
# Load EWS Managed API
if (!(LoadEWSManagedAPI))
{
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}

# Check we have valid credentials
if ($Credentials -ne $Null)
{
    If ($Username -or $Password)
    {
        Write-Host "Please specify *either* -Credentials *or* -Username and -Password" Red
        Exit
    }
}

#$LogExists = Test-Path $LogFile
#if ($LogExists){Remove-Item  $LogFile}

Write-Host ""

# Check whether we have a CSV file as input...
$DataFile = "Users.csv"
$FileExists = Test-Path $DataFile

If ( $FileExists )
{
	# We have a CSV to process
    Write-Host "Reading mailboxes from CSV file"
	$csv = Import-CSV $DataFile -Header "PrimarySmtpAddress"
	foreach ($entry in $csv)
	{
        Write-Verbose $entry.PrimarySmtpAddress
        if (![String]::IsNullOrEmpty($entry.PrimarySmtpAddress))
        {
            if (!$entry.PrimarySmtpAddress.ToLower().Equals("primarysmtpaddress"))
            {
		        $Mailbox = $entry.PrimarySmtpAddress
			    SearchSentItems $Mailbox
            }
        }
	}
}
Else
{
	# Process as single mailbox
	Write-Host "Did not find the CSV file to read the mailboxes from."
}
