#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
# The script requires the EWS managed API, which can be downloaded here:
# http://www.microsoft.com/en-us/download/details.aspx?id=35371
# This also requires PowerShell 2.0
# Make sure the Import-Module command below matches the DLL location of the API.
# This path must match the install location of the EWS managed API. Change it if needed.

#To View Only
#.\CheckContainerClassV2.ps1 -Username ServiceAccount -Password PSSLab1! -Domain EXHR-0512Dom -EwsUrl "https://EXHR-0512.EXHR-0512dom.extest.microsoft.com/ews/exchange.asmx" -IgnoreSSLCertificate -Impersonate

param (
	[Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")]
	[string]$Username,
	
	[Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")]
	[string]$Password,
	
	[Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")]
	[string]$Domain,
	
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
	
	[Parameter(Mandatory=$False,HelpMessage="Just Show the Container Class")]	
	[switch]$Archive,
	
	[Parameter(Mandatory=$False,HelpMessage="Log file - activity is logged to this file if specified")]	
	[string]$LogFile = ""
	
)

[string]$info = "White"                # Color for informational messages
[string]$warning = "Yellow"            # Color for warning messages
[string]$error = "Red"                # Color for error messages
[string]$success = "Green"                # Color for success messages


[string]$LogFile = "CheckContainerClassv2Log.txt"   # Path of the Log File

$verbose = $true

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
    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)

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

    return $exchangeService
}

function ProcessMailbox($MailboxName)
{
		# Process the mailbox
		Log "$([string]::Format('Processing mailbox {0}', $MailboxName))" $info 
		
		$global:service = CreateService($MailboxName)
		if ($Null -eq $global:service)
		{
			Write-Host "Failed to create ExchangeService" -ForegroundColor $error
		}
	
	
	try
	{
		
	#Number of Items to Get
    $FpageSize =50
    $FOffset = 0
    $MoreFolders =$true

    while ($MoreFolders)
     {

        #Setup the View to get a limited number of Items at one time
        $folderView = new-object Microsoft.Exchange.WebServices.Data.FolderView($FpageSize,$FOffset,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
        $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        
        # Folder Path
        $PR_FOLDER_PATH = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
		$PR_CONTAINER_CLASS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3613, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
		

        $FPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet(
                            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
                            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
                            [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass);
                            
        $FPropertySet.Add($PR_FOLDER_PATH);  
        $FPropertySet.Add($PR_CONTAINER_CLASS);  
		
        $folderView.PropertySet = $FPropertySet;
        
        #Create the Search Filter.
		$FoSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_CONTAINER_CLASS, "")
		
		if ($Archive)
			{
				$oFindFolders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$FoSearchFilter,$folderView)
			}
			else{
		        $oFindFolders = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$FoSearchFilter,$folderView)

			}
		
        
        foreach ($folder in $oFindFolders.Folders)
        {
	       
            &{
				Log "Folder Path: '$($folder.ExtendedProperties[0].Value)', PR_CONTAINER_CLASS: '$($folder.ExtendedProperties[1].Value)'." $info
            }
            trap [System.Exception] 
            {
                Write-host ("Error: " + $_.Exception.Message) -foregroundcolor $error;
                continue;
            }
        }

         if ($oFindFolders.MoreAvailable -eq $false)
            {$MoreFolders = $false}

             if ($MoreFolders)
            {$FOffset += $FpageSize}
        
     }		
	
	}
    catch
    {
        Write-host ("Error: Either the folder already exists OR something else went wrong.") -foregroundcolor $error;
   		Continue
    }		
   
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
if ($Null -ne $Credentials)
{
    If ($Username -or $Password)
    {
        Write-Host "Please specify *either* -Credentials *or* -Username and -Password" Red
        Exit
    }
}

$LogExists = Test-Path $LogFile
if ($LogExists){Remove-Item  $LogFile}

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
		        #$Mailbox = $entry.PrimarySmtpAddress
			    ProcessMailbox($entry.PrimarySmtpAddress)
            }
        }
	}
}
Else
{
	# Process as single mailbox
	Write-Host "Did not find the CSV file to read the mailboxes from."
}