
$dictionaryPath =  (split-path $myinvocation.mycommand.path) + "\dict.txt"
if (test-path $dictionaryPath) {
	$dictionary = get-content $dictionaryPath
} else {
	throw "Dictionary file for passphrase generation '$dictionaryPath' could not be found."
}

$exchangePowershellConnectionURI = "http://exchangeserver.domain.local/powershell"
$domain = "domain" # use domain short name
$emaildomain = "domain.org"


$date = get-date -format "yyyy-MM-dd-hh-mm"
$logFilePath = $date + "-Account Creation.txt"
$csvFilePath = $date + "-import for KeePass.csv"
$iCreateDB7 = "\\domain.local\applications\Esquire\iCreate7\iCreate.mdb"
$iCreateDB7EP2 = "\\domain.local\applications\Esquire\iCreate7ep2\iCreate.mdb"

$userPrincipalNameSuffix = "@domain.org"
$userEmailSuffix = $userPrincipalNameSuffix

################################################
# Variables that control where shared / private folders are created for each new user
$userShare = "\\domain.local\Shares\Users\"
$Site2Share = "\\domain.local\Shares\SharedSite2\"
$Site1Share = "\\domain.local\Shares\SharedSite1\"
$Site2Drive = "S:\"
$Site1Drive = "Q:\"

#############################################################
# Variables for sending email messages to new staff
$emailFrom = "helpdesk@domain.org"
$emailServer = "smtp.domain.org"
# This is the location of emails. Emails in the _Everyone subfolder will be sent to all departments. The Subject is taken from the
# name of the file, the body from the actual file contents. File should be HTML formatted.
$emailsPath = "\\domain.local\shares\SharedSite2\Administration\IT\Scripts, Useful for Administration\User Provisioning\Emails"

#############################################################
# Variables for updating Office 365 licenses
$Office365User = "useraccountcreation@domainma.onmicrosoft.com"
$Office365Password = "xxxxxxxxxxxxxxxxxxxxx"
$Office365DirSyncComputer = "ADFS-HOST.domain.local"
$Office365AccountSkuId = "domain:STANDARDWOFFPACK"
$Office365UsageLocation = "US"
$Office365TenantAddressSuffix = "@domain.mail.onmicrosoft.com"
$MimecastMSEServiceAccount = "servicemimecastmse@domain.org" # used for granting permissions for MSE / folder view in Mimecast
$Office365MySitesBaseURL = "https://my.domain.org/Person.aspx?accountname=domain\"

#############################################################
# Information used to update the new account database
$pathToKPScript = "C:\Program Files (x86)\KeePass Password Safe 2\kpscript.exe"
$pathToKeePassDB = "\\domain\shares\SharedSite2\Administration\Personnel\New User Accounts\User Account Information.kdbx"

#############################################################
# Variables for adding accounts to Legal Server
# Begin block (specify on command line)
$lsLoginURL = "https://CONTOSO.legalserver.org"
$lsNewUserURL = "https://CONTOSO.legalserver.org/user/process/dynamic_87/"
$lsUsername = "useraccountcreation"
$lsAdminPassword = "XXXXXXXXXXXXXXXXXXXXXXX"
# Process block (should be specified in $global:usersummary psobject which is defined in accountcreation.psm1)
$lsTypeIDStaff = "501"
$lsTypeIDVolunteer = "260130"
$lsRoleIDUser = "844"
$lsRoleIDVolunteer = "843"
$lsOfficeIDMainOffice = "1"

##############################
# Information specific to each department
# 	OU is the location in AD where the new account should be placed. If left empty, it is inherited from template account
#	Folder is the subfolder where non-student shared folders are created
#	students is the subfolder where student shared folders are created
#	template is the SamAccountName of the account that will be used as a template (including group membership)
$global:depts =@{
			"Administration" =  
				@{	"OU" = "OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"folder" = "Administration\";
					"students" = "Administration\";
					"template" = "tadministration";
					"LSProgramID" = "257739"};
			"Accounting" =  
				@{	"folder" = "Administration\Accounting\";
					"OU" = "OU=Accounting,OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"students" = "Administration\Accounting\";
					"template" = "tadministration";
					"LSProgramID" = "257739"};
			"Development" =
				@{	"folder" = "Administration\Development\";
					"OU" = "OU=Development,OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"students" = "Administration\Development\";
					"template" = "tdevelopment";
					"LSProgramID" = "257739"};
			"Facilities" =  
				@{	"folder" = "Administration\Facilities\";
					"OU" = "OU=Facilities,OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"students" = "Administration\Facilities\";
					"template" = "tadministration";
					"LSProgramID" = "257739"};	
			"IT" =  
				@{	"folder" = "Administration\IT\";
					"ou" = "";
					"students" = "Administration\IT\";
					"template" = "tit";
					"LSProgramID" = "257739"};
			"Management" =  
				@{	"folder" = "Administration\Management\";
					"OU" = "OU=Management,OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"students" = "Administration\Management\";
					"template" = "tadministration";
					"LSProgramID" = "257739"};	
			"Payroll" =  
				@{	"folder" = "Administration\Payroll\";
					"OU" = "OU=Payroll,OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"students" = "Administration\Payroll\";
					"template" = "tadministration";
					"LSProgramID" = "257739"};
			"Personnel" =  
				@{	"folder" = "Administration\Personnel\";
					"OU" = "OU=Personnel,OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"students" = "Administration\Personnel\";
					"template" = "tpersonnel";
					"LSProgramID" = "257739"};
			"Reception" =  
				@{	"folder" = "Administration\Reception\";
					"OU" = "OU=Reception,OU=Administration,OU=Boston,OU=_Domain_users,DC=GBLS,DC=Local";
					"students" = "Administration\Reception\";
					"template" = "treception";
					"LSProgramID" = "257739"};	
}