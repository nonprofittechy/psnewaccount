
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


##############################
# Information specific to each department
# 	OU is the location in AD where the new account should be placed. If left empty, it is inherited from template account
#	Folder is the subfolder where non-student shared folders are created
#	students is the subfolder where student shared folders are created
#	template is the SamAccountName of the account that will be used as a template (including group membership)
$global:depts =@{
			"Administration" =  
				@{	"OU" = "OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"folder" = "Administration\";
					"students" = "Administration\";
					"template" = "tadministration"};
			"Accounting" =  
				@{	"folder" = "Administration\Accounting\";
					"OU" = "OU=Accounting,OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"students" = "Administration\Accounting\";
					"template" = "tadministration"};
			"Development" =
				@{	"folder" = "Administration\Development\";
					"OU" = "OU=Development,OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"students" = "Administration\Development\";
					"template" = "tdevelopment"};
			"Facilities" =  
				@{	"folder" = "Administration\Facilities\";
					"OU" = "OU=Facilities,OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"students" = "Administration\Facilities\";
					"template" = "tadministration"};	
			"IT" =  
				@{	"folder" = "Administration\IT\";
					"ou" = "";
					"students" = "Administration\IT\";
					"template" = "tit"};
			"Management" =  
				@{	"folder" = "Administration\Management\";
					"OU" = "OU=Management,OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"students" = "Administration\Management\";
					"template" = "tadministration"};	
			"Payroll" =  
				@{	"folder" = "Administration\Payroll\";
					"OU" = "OU=Payroll,OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"students" = "Administration\Payroll\";
					"template" = "tadministration"};
			"Personnel" =  
				@{	"folder" = "Administration\Personnel\";
					"OU" = "OU=Personnel,OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"students" = "Administration\Personnel\";
					"template" = "tpersonnel"};
			"Reception" =  
				@{	"folder" = "Administration\Reception\";
					"OU" = "OU=Reception,OU=Administration,OU=Site2,OU=_Domain_users,DC=domain,DC=Local";
					"students" = "Administration\Reception\";
					"template" = "treception"}			

}