##############################################################
#
#				Provision User Accounts - provision.ps1
#				Purpose: using a .csv file as source, create user accounts	
#				Enables mailbox, creates home folders, and adds
#				to iCreate DBs
#				Will also generate unique, long passwords
#
##############################################################


##############################################################
#
#				Initialize Variables
#				Note: these variables must be updated first if anything changes, including script location
#
##############################################################

# Use implicit remoting to connect to Exchange Server for Exchange cmdlets (and account creation)
# import the provision.variables.ps1 file
$variablesPath = Join-path (split-path $myinvocation.mycommand.path)  provision.variables.ps1
. $variablesPath

$session = new-pssession -configurationname Microsoft.exchange -connectionuri $exchangePowershellConnectionURI
import-pssession -DisableNameChecking $session | out-null
import-module -DisableNameChecking activedirectory | out-null

$global:userSummary= @()

#####################################################
#
# 				Main Function
#
######################################################

function Provision {
	PROCESS {
		$user = CreateUser $_ 
		if (test-user($user.samAccountName)) {
			CreateHomeFolder $user
			# iCreate database editing needs to run in 32 bit process for opening .MDB file
			start-job -scriptblock $AddToiCreateDB -argumentList $user,$icreateDB7 -runas32 | wait-job | receive-job 
		}
		VerifyAccount $user
	}
}

# outputs user account data to the pipeline, takes input in the form of a CSV file
function ProvisionInputCSV($filename) {

	$users = Import-CSV $filename
	
	foreach ($user in $users) {
		
		
		$ht = @{'givenName'=	$user."First Name".trimEnd();
				'sn'=			$user."Last Name".trimEnd();
				'mi' =			$user."Middle Initial".trimEnd();
				'phone' = 		$user."Telephone".trimEnd();
				'title'=		$user.Title.trimEnd();
				'department'=	$user.Unit.trimEnd();
				'manager'=		$user.Manager.trimEnd()
				}

		# check to see if account expiration date was set and is a valid date in the future
		# if so, cast it to a date and set to accountExpirationDate
		$ed = $user."Account Expiration Date".trimEnd()
		$today = [DateTime]::Today
		if ( ($ed -as [DateTime] -ne $null) -and (($ed -as [DateTime]) -ge $today)) {
			$ht.accountExpirationDate = [DateTime]$ed
		} else {
			$ht.accountExpirationDate = $null
			OutputLog("Account expiration date for " + $ht.givenName + " " + $ht.sn + " was unspecified, invalid, or is in the past. Ignoring!")
		}
				
		# The next steps rely on valid input data--we won't run them if any necessary fields are empty
		if ($ht.sn -ne "" -and $ht.givenName -ne "" -and $ht.title -ne "" -and (isValidDepartment($ht.department))) {
		
			# Create generated attributes for user account
			$ht.ClearTextPassphrase = passphrase($dictionary)
			$ht.samAccountName = (UniqueSamAccountName($ht))
			$ht.name = (UniqueName($ht))
			$ht.userSharedFolder = (getUserSharedFolderPath($ht))
			$ht.userSharedFolderOther = (getUserSharedFolderPathOther($ht))
			$ht.manager = getManagerSAMAccountName($ht.manager);
			
			# send the cleaned-up data to be processed
			Write-Output $ht
		} else {
			OutputLog("Detected an empty line or missing data. Please make sure all accounts you requested were created.")
		}
	}
}

##################################################################
#
#				Account Creation Sub-Functions
#
##################################################################


function isValidDepartment($department) {
	return $depts.contains($department)
}

function isValidManager($manager) {
	return $manager -ne ""
}

function getManagerSAMAccountName($manager){
	if ($manager.trimEnd() -eq "") {
		OutputLog("*** WARNING *** : Manager name not provided.")
		return ""
	}
	if (test-user($manager)) {
		return $manager
	} else {
		# Let's see if the name was mistakenly given in the format "First Last" or "Last First"
		$names = $names.split("\s+")
		$firstlast = get-aduser -filter {sn -like $names[1] -and givenname -like $names[0]}
		if ($firstlast -ne $null -and !($firstlast -is [array]) ) {
			return $firstlast.samAccountName
		}
		$lastfirst = get-aduser -filter {sn -like $names[0] -and givenname -like $names[1]}
		if ($lastfirst -ne $null -and !($lastfirst -is [array]) ) {
			return $lastfirst.samAccountName
		}
		OutputLog("*** Warning: Manager name " + $manager + " appears to be ambiguous or invalid. Please provide a user id, such as jbowman. Will try using the template's manager instead.")
		return ""
	}		
}

##########################
# Main function: 
# Create user account and remote mailbox on Office 365
function CreateUser($userinfo) {
	# we will select a domain controller here and use it for both creating the AD account and mail-enabling it, to prevent race condition
	$DC = Get-adDomainController | select -expand "HostName"
		
	$userinfo.upn = ($userinfo.samAccountName + $userPrincipalNameSuffix)
	$securePassphrase = convertto-securestring -asplaintext -force -string $userinfo.ClearTextPassphrase
	$template = get-aduser -identity $depts[$userinfo.department].template -properties city,country,company,department,description,Fax,HomePage,Manager,MemberOf,Organization,PostalCode,ScriptPath,State,StreetAddress,telephoneNumber,wWWHomePage
	
	# was the OU explicitly set? If not, inherit from template
	if (!($depts[$userinfo["department"]].OU -eq "")) {
		$userinfo.ou = $depts[$userinfo["department"]].OU
	} else {
		if ($template.distinguishedName -match ".*?(OU=.*)") {
			$userinfo.ou = $matches[1]
		}
	}
	
	if ($userinfo.mi.length -gt 0) {
		$mi = $userinfo.mi.substring(0,1)
	} else {
		$mi = ""
	}
	
	$userinfo.homeDirectory = ($userShare + $userinfo.samAccountName)
	$userinfo.emailAddress = ($userinfo.samAccountName + $userEmailSuffix)
	
	# check the manager field -- earlier we set to the empty string if it didn't look valid. If it's empty, use the template.
	if (isValidManager($userinfo.manager)) {
		$manager = $userinfo.manager
	} else {
		$manager = $template.manager
		OutputLog("*** WARNING *** Manager was not specified or was invalid for user " + $userinfo.sn + ", " + $userinfo.givenName + ". Using the unit template's manager, " + $manager + ".")
	}
	
	$null = New-ADUser -instance $template `
		-name $userinfo.samAccountName `
		-ChangePasswordAtLogon $true `
		-DisplayName ($userinfo.sn + ", " + $userinfo.givenName) `
		–GivenName 	$userinfo.givenName `
		-Surname 	$userinfo.sn `
		-Initials 	($userinfo.givenName.substring(0,1) + $mi + $userinfo.sn.substring(0,1)) `
		-OfficePhone $userinfo.phone `
		-Title $userinfo.title `
		-department $userinfo.department `
		-enabled $true `
		-emailaddress  $userinfo.emailAddress `
		-UserPrincipalName ($userinfo.upn) `
		-path $userinfo.ou `
		-Manager $manager `
		-Description ($userinfo.title + ": " + $userinfo.department) `
		-AccountPassword $securePassphrase `
		-accountExpirationDate ($userinfo.accountExpirationDate) `
		-HomePage ($Office365MySitesBaseURL + $userinfo.samAccountName) `
		-Server $DC
	
	#	Wait up to 30 seconds for account creation to be complete
	# 	Use the same DC as above to reduce risk of race condition
	$i =  30
	while( ($i -gt 0) -and !(test-user -samAccountName $userinfo.samaccountname -adDomainControllerHostName $DC) ) {
		write-host "Waiting up to $i seconds for account creation..."
		$i--
		sleep 1
	}

	##########################################
	#	Perform all steps that can't be taken when account is first created
	# 	and those dependent on being a student/full user
	# 	Use the same DC as above to reduce risk of race condition
	if ( test-user -samAccountName $userinfo.samaccountname -adDomainControllerHostName $DC ) {
	
		$userSharedFolder = [string]$userinfo.userSharedFolder
		$userSharedFolderOther = [string]$userinfo.userSharedFolderOther
	
		#work around an odd quirk in AD powershell to allow us to use a comma in the account name
		get-aduser -identity $userinfo.samAccountName | rename-adobject -newname $userinfo.name | out-null
	
		# map the home drive to u: for all users, except interns
		# also, map the attribute "userSharedFolder" to the user's shared directory on s: drive
		# and add interns to the "Interns" group.
		if ( ($userinfo.title -match "Intern" ) -or ($userinfo.title -match "Student") ) {
			set-aduser -identity $userinfo.samAccountName -replace @{userSharedFolder=$userSharedFolder;userSharedFolderOther=$userSharedFolderOther}
			$null = Add-ADGroupMember -identity "Interns" -members $userinfo.samAccountName
		} else {
			set-aduser -identity $userinfo.samAccountName `
					   -HomeDirectory ($userinfo.homeDirectory) `
				       -homeDrive U: `
				       -replace @{userSharedFolder=$userSharedFolder;userSharedFolderOther=$userSharedFolderOther}
		}
			
		# add the groups from the template object
		foreach($group in $template.memberOf) {
			$null = Add-ADGroupMember -identity $group -members $userinfo.samAccountName
		}
		
		# Mail-enable the account on Office 365
		# Use the same domain controller we used to create the account
		$null = new-office365mailbox -userinfo $userinfo -DomainController $DC
					
		return $userinfo
	
	} else {
		OutputLog("Warning: user account " + $userinfo.name + " (" + $userinfo.samAccountName + ")" + " doesn't seem to have been created successfully yet, and will not be mail enabled. Account creation may be taking an abnormally long amount of time. If the account was created, try manually mail-enabling it.")
	}
}

#########################
# Create home folder / unit subfolder
function CreateHomeFolder($userinfo) {
	$path = ""
	
	# Create the Unit folder for the user (on S: drive)
	if ($userinfo["department"] -eq "Site1") {
		$path = $Site1Share
	} else {
		$path = $Site2Share
	}
	
	# create a U: drive folder only for full-time staff, not students
	if (($userinfo["title"] -match "Intern" ) -or ($userinfo["title"] -match "Student")) {
		$path += $depts[$userinfo["department"]].students + $userinfo["samAccountName"]
	} else {
		# Create the Users share (U: drive)
		.\setfolderpermission.ps1 -path ($userShare + $userinfo.samAccountName) -access ($domain+"\"+$userinfo.samAccountName) -permission FullControl
		$path += $depts[$userinfo["department"]].folder + $userinfo["samAccountName"]
	}
	
	if (test-path ($path) ) {
		OutputLog("*** WARNING *** User's Unit folder $path already exists! Skipping creating it again. Verify permissions are correct.") 
	} else {
		# Create the path but just use the inherited permissions--no need to set special permissions for this folder
		mkdir $path
	}
}

#########################
# Add to the iCreate database
# Create as an object to allow invoking as a script block -- simplifies calling the 32-bit version of the ODBC driver from 64 bit Powershell

$addtoiCreateDB = {

	param ($userInfo, $icreateDB7)
	
	write-host "Checking to see if we should add user to iCreate ..."
#	process {
			#function addtoiCreateDB($userInfo) { # Remove to allow invoking as start-job

		switch -regex ($userinfo["Title"]) {
			".*Attorney.*" {$userinfo.employeeType = 1}
			".*Paralegal.*" {$userinfo.employeeType = 2}
			".*((Student)|(Intern))" {return 0} # don't add students to the iCreate DB
			Default {$userinfo.employeeType = 3}
		}
		
		if( !(test-path -path $iCreateDB7) ) { 
			throw ($icreateDB7 + " does not exist!") 
		}
			
		#format the icreate-specific input variables
		if($userinfo['department'] -eq "Site1") {
			$userinfo.officeID = 2
			$userinfo.generalNo="(555) 555-5555"
			$userinfo.faxNo = "(555) 555-5555"
		} else {
			$userinfo.officeID = 1
			$userinfo.generalNo = "(555) 555-5555"
			$userinfo.faxNo = "(555) 555-5555"
		}
		
		if(!($userinfo["mi"] -eq "")) {
			$userinfo.middleInitial = $userinfo["mi"] + ". "
		} else {$userinfo.middleInitial = ""}
		
		if ($userinfo.phone.length -gt 0) {
			$userinfo.directNoExt = $userInfo["phone"].SubString($userinfo["phone"].length-4)
		} else {
			$userinfo.directNoExt = "1234"
		}
		
		$dbobj = New-Object -ComObject ADODB.Connection
		
		if (test-path $icreateDB7) {
			if ($userinfo.givenName -ne "" -and $userinfo.sn -ne "" -and $userinfo.samAccountName -ne "") {
					$dbObj.Open("Provider = Microsoft.Jet.OLEDB.4.0;Data Source=$iCreateDB7" ) 
					$recordSet = New-Object -ComObject "ADODB.Recordset"
					
					$adOpenStatic = 3
					$adLockOptimistic = 3
				
					$recordset.Open("Select * From Authors", $dbObj,$adOpenStatic,$adLockOptimistic)
				
					$recordSet.AddNew()
					
					$recordSet.Fields.Item("FirstName").Value = $userinfo["givenName"]
					$recordSet.Fields.Item("MiddleName").Value = $userinfo["mi"]
					$recordSet.Fields.Item("LastName").Value = $userinfo["sn"]
					$recordSet.Fields.Item("OfficeId").Value = $userinfo.officeID
					$recordSet.Fields.Item("Department").Value = $userinfo["department"]
					$recordSet.Fields.Item("Closing").Value = "Sincerely,"
					$recordSet.Fields.Item("ClosingName").Value = $userinfo["givenName"] +" " + $userinfo.middleInitial + " " + $userinfo["sn"]
					$recordSet.Fields.Item("FontPreference").Value = "Times New Roman"
					$recordSet.Fields.Item("DirectNo").Value = $userinfo["phone"]
					$recordSet.Fields.Item("DirectNoExt").Value = $userinfo.directNoExt
					$recordSet.Fields.Item("GeneralNo").Value = $userinfo.GeneralNo
					$recordSet.Fields.Item("FaxNo").Value = $userinfo.FaxNo
					$recordSet.Fields.Item("EmployeeType").Value = $userinfo.employeeType
					$recordSet.Fields.Item("Title").Value = $userinfo["title"]
					$recordSet.Fields.Item("Email").Value = $userinfo["emailAddress"]
					$recordSet.Fields.Item("userID").Value = $userinfo.samAccountName

					$recordSet.Update()
					$dbObj.close()
				} else {
					write-host "*** Warning *** not adding to the iCreate database as we seem to have been given an empty record to create."
					#OutputLog("*** Warning *** not adding to the iCreate database as we seem to have been given an empty record to create.")
				}
			} else {
				
				write-host "*** WARNING *** Couldn't add to iCreate Database: iCreate database missing: $iCreateDB7"
		}	
#	}
}

#################################################
# Get a path name for the user's shared folder to pass to the folder creation function
# based on user's department, site, and intern/full time staff
function getUserSharedFolderPath($ht) {

		if ($ht.department -eq "Site1") {
			$userSharedFolder = $Site1Share
		} else {
			$userSharedFolder = $Site2Share
		}
		
		if (($ht.title -match "Intern" ) -or ($ht.title -match "Student")) {
			$userSharedFolder += $depts[$ht.department].students + $ht.samAccountName
		} else {
			$userSharedFolder += $depts[$ht.department].folder + $ht.samAccountName
		}
		
		return [string]$userSharedFolder
}

function getUserSharedFolderPathOther($ht) {

		if ($ht.department -eq "Site1") {
			$userSharedFolderOther = $Site1Drive
		} else {
			$userSharedFolderOther = $Site2Drive
		}
		
		if (($ht.title -match "Intern" ) -or ($ht.title -match "Student")) {
			$userSharedFolderOther += $depts[$ht.department].students + $ht.samAccountName
		} else {
			$userSharedFolderOther += $depts[$ht.department].folder + $ht.samAccountName
		}
		
		return [string]$userSharedFolderOther
}

####################################################
# 	Verify account creation
# 	output to the log file and add to the array for CSV to import into KeePass / sending of emails
function VerifyAccount($user) {
	if (!(test-user($user.samAccountName) )) {
		OutputLog("*** WARNING *** User's Active Directory account "  + $user.name + " (" + $user.samAccountName + ")" + " appears not to have been created. Please wait a few minutes and verify account creation. Mailbox may need to be created manually.")
	} else {
		OutputLog("Successfully created new user in Active Directory: " + $user.name + ", login name: " + $user.samAccountName + ", Department: " + $user.department )
		
		$userObj = @{"Title" = $user.name; 
				   "Username" = $user.samAccountName;
				   "Password" = $user.clearTextPassphrase;
				   "Url" = "";
				   "Notes" = ("Unit: " + $user.Unit + "`r`n Shared Folder: " + $user.userSharedFolder);
				   "samAccountName" = $user.samAccountName;
				   "Department" = $user.department;
				   "emailAddress" = $user.emailAddress;
				   "role" = $user.title;
				   "accountExpirationDate" = $user.accountExpirationDate;
				   "SharedFolder" = $user.userSharedFolderOther
				  }
	
		$global:userSummary += new-object psobject -property $userObj
			
	}
	
	if (!(test-mailboxenabled($user.samAccountName))) {
		OutputLog("*** WARNING *** User account " + $user.name + "(" + $user.samAccountName + ")" + " couldn't be found. The account may not be mail-enabled, or there may just be a lag in syncing between domain controllers.")
	} else {
		OutputLog("Successfully requested an Exchange Mailbox for "  + $user.name + " (" + $user.samAccountName + ") Mailbox should be created at the next synchronization with Office 365 (every 3 hours).")
	}
}

################################################################
#
#					Library / reusable Functions
#
###############################################################

# Optional: specify the hostname of the domain controller you want to check
function test-user($samAccountName, $adDomainControllerHostName) {
	if ($adDomainControllerHostName -eq $null) {
		$adDomainControllerHostName = get-addomaincontroller | select -expand "HostName"
	}
	return [bool](Get-AdUser -server $adDomainControllerHostName -filter {samaccountname -like $samAccountName})
}

function test-mailboxenabled($samAccountName) {
	$u = get-aduser $samaccountname -properties Mail
	
	return ! ([bool]($u.Mail -eq $null))
}

#Given a dictionary, generate a random 2-word passphrase. The first letter is capitalized and a random number is appended to meet
# complexity requirements
function PassPhrase ($words) {
	$phrase = get-random $words
	$phrase += " "
	$phrase += get-random $words
	$phrase = [string]$phrase	
	$phrase = ( ($phrase.subString(0,1).toUpper()) + $phrase.substring(1) + (get-random -max 99))	
	return [string]$phrase
}

Function OutputLog($textOrObject) {
	$textOrObject >> $logFilePath
}

##############################
# Generate a unique account name given user info
# Try FLastname, followed by FMLastname, followed by FirstLastname, followed by FirstMLastname, followed by FL#
# strip out any non alpha-numeric characters along the way and any leading digits
Function UniqueSamAccountName($userinfo) {	

	$gn = ($userinfo.givenName -replace  '[^a-zA-Z0-9]','') -replace '^\d+'
	$sn = ($userinfo.sn -replace '[^a-zA-Z0-9]','') -replace '^\d+'
	$mi = ($userinfo.mi -replace  '[^a-zA-Z0-9]','') -replace '^\d+'
	
	$samAccountName = ($userinfo.givenName.substring(0,1) + $sn)
	if ((Get-ADUser -filter {samaccountname -like $samAccountName}) ) {
		$samAccountName = ($gn.substring(0,1) + $mi + $sn)
		if ( (Get-ADUser -filter {samaccountname -like $samAccountName}) ){
			$samAccountName = ($gn + $sn)
			if ( (Get-ADUser -filter {samaccountname -like $samAccountName}) ){
				$samAccountName = ($gn + $mi + $sn)	 		
				$i = 2 #  Logically it's at least the second user with this name--avoid confusion, start numbering at 2!
				while ((Get-ADUser -filter {samaccountname -like $samAccountName} )) {
					$samAccountName = ($gn.substring(0,1) + ($sn.substring(0,1)) + $i)
					$i++
				}				
			}			
		}		
	}	
	return $samAccountName
}

###########################
# Exchange requires a unique account name, separate from unique samAccountName and UPN :(
# too bad for folks born with common names.
# First, set to Lastname, firstname, then try Lastname, firstname (mi), then try Lastname, 
# firstname (userID), finally fall back on adding a sequential number to the username (should never be necessary though)
function UniqueName($userinfo) {
	$Name = ($userinfo.sn + ", " + $userinfo.givenName)
	
	if (Get-ADUser -filter {name -like $name}) {
		if ($userinfo.mi -ne "") {
			$name += " " + $userinfo.mi
		}
		$i = 2
		if (get-ADUser -filter {name -like $name} ) {
			$name = ($userinfo.sn + ", " + $userinfo.givenName + " (" + $userinfo.samAccountName + ")")
		}
		while (Get-ADUser -filter {name -like $name}) {
			$name = ($userinfo.sn + ", " + $userinfo.givenName) + $i
			$i++
		}
		OutputLog("*** WARNING *** User name "+ $userinfo.sn + ", " + $userinfo.givenName + " is already in use for a different user. Instead, it was changed to $name. Please double-check with the user to verify preferred name.")
		return $name
	} else {
		return $name
	}
}

# http://blogs.technet.com/b/heyscriptingguy/archive/2009/09/01/hey-scripting-guy-september-1.aspx
Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "CSV Files (*.csv)| *.csv"
 $OpenFileDialog.ShowHelp = $true
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
 #$OpenFileDialog.ShowHelp = $true
} #end function Get-FileName

############################################
# Office 365 functions

# creates a new mailbox with the tenant address set in provision.variables.ps1
# See: http://www.office365tipoftheday.com/2014/01/31/remote-mailboxes/
function new-office365mailbox ($userinfo, $domaincontroller) {
	$routingAddress = $userinfo.samAccountName + $Office365TenantAddressSuffix
	enable-remotemailbox ($userinfo.upn) -primarysmtpaddress ($userinfo.emailAddress) -remoteroutingaddress $routingaddress -domaincontroller $domaincontroller
	set-remotemailbox -identity ($userinfo.upn) -emailAddresses @{add=$routingAddress} -domaincontroller $domaincontroller
}

# Export the user summary we've been building to a CSV file
Function export-keepassCSV ($path) {
	$global:userSummary | select-object Username,Title,Url,Password,Notes | export-csv -path $path -noTypeInformation
}

#########################################################################################
#
#					Action
#########################################################################################

$csvFile = get-FileName("\\Domain\Shares\SharedSite2")

if(!($csvFile -eq "" ) -and (test-path $csvFile)) {
	provisionInputCSV($csvFile) | provision
	export-keepasscsv -path $csvFilePath
	
	write-host "`r`n`r`nCreated " ($global:userSummary.length) " new users:"
	$global:userSummary | select username, emailaddress, role, department, SharedFolder | format-table
	
	write-host "View $logFilePath for summary of results. `r`n Import $csvFilePath into KeePass to share usernames and passwords with Personnel. `r`n Please delete the CSV file immediately after import."
}