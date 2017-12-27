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

# Use implicit remoting to connect to Server for Exchange cmdlets (and account creation)
# import the provision.variables.ps1 file
$cwd = split-path $myinvocation.mycommand.path
$variablesPath = Join-path $cwd  provision.variables.ps1
. $variablesPath

$session = new-pssession -configurationname Microsoft.exchange -connectionuri $exchangePowershellConnectionURI
import-pssession -DisableNameChecking $session | out-null
import-module -DisableNameChecking activedirectory | out-null

import-module .\accountcreation.psm1
. (join-path $cwd "folderhandling.ps1")
. (join-path $cwd "icreate.ps1")
. (join-path $cwd "office365.ps1")

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
			start-job -scriptblock $AddToiCreateDB -argumentList $user,$icreateDB7EP2 -runas32 | wait-job | receive-job 
		}
		$global:userSummary += (VerifyAccount -user $user)
	}
}

# outputs user account data to the pipeline, takes input in the form of a CSV file
function ProvisionInputCSV($filename) {

	$DC = Get-adDomainController | select -expand "HostName"

	$users = Import-CSV $filename
	
	foreach ($user in $users) {
		
    $origUnit = $user.Unit
		# do some input cleanup for the more complicated unit names
		switch -regex ($user.Unit) {
			"Consumer" {$user.Unit = "Consumer Rights Project"}
#			"(Elder)|(Health)" {$user.Unit = "Elder, Health, Disability"}
			"(AOU)|(Asian)" {$user.Unit = "Asian Outreach Unit"}
			"Cambridge" {$user.Unit = "CASLS"}
			"University" {$user.Unit = "BU"} 
		}
		
		$ht = @{'givenName'=	$user."First Name".trimEnd();
				'sn'=			$user."Last Name".trimEnd();
				'mi' =			$user."Middle Initial".trimEnd();
				'phone' = 		$user."Telephone".trimEnd();
				'title'=		$user.Title.trimEnd();
				'department'=	$user.Unit.trimEnd();
				'manager'=		$user.Manager.trimEnd();
        'origUnit'=   $user.origUnit
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
			$ht.manager = getManagerSAMAccountName -manager $ht.manager -adDomainControllerHostName $DC;
			
			# send the cleaned-up data to be processed
			Write-Output $ht
		} else {
			OutputLog("Detected an empty line or missing data. Please make sure all accounts you requested were created.")
		}
	}
}


##########################
# Main function: 
# Create user account and remote mailbox on Office 365
function CreateUser($userinfo) {
	# we will select a domain controller here and use it for both creating the AD account and all subsequent operations, to prevent race condition
	$DC = Get-adDomainController | select -expand "HostName"
		
	$userinfo.upn = ($userinfo.samAccountName + $userPrincipalNameSuffix)
	$securePassphrase = convertto-securestring -asplaintext -force -string $userinfo.ClearTextPassphrase
	$template = get-aduser -server $DC -identity $depts[$userinfo.department].template -properties city,country,company,department,description,Fax,HomePage,Manager,MemberOf,Organization,PostalCode,ScriptPath,State,StreetAddress,telephoneNumber,wWWHomePage
	
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
		-samAccountName $userinfo.samAccountName `
		-name $userinfo.name `
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
		-HomePage ($Office365MySitesBaseURL + $userinfo.upn) `
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
		# get-aduser -identity $userinfo.samAccountName | rename-adobject -newname $userinfo.name | out-null
	
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

#########################################################################################
#
#					Action
#########################################################################################

$csvFile = get-FileName(".\")

if(!($csvFile -eq "" ) -and (test-path $csvFile)) {
	provisionInputCSV($csvFile) | provision
	export-keepasscsv -path $csvFilePath
  
  # Add a user account into Legal Server
  #$global:userSummary | new-legalserveraccount -loginURL $lsLoginURL -newUserURL $lsNewUserURL -adminUsername $lsUsername -adminPassword $lsAdminPassword
	
	write-host "`r`n`r`nCreated " ($global:userSummary.length) " new users:"
	$global:userSummary | select username, emailaddress, role, department, SharedFolder | format-table
	
	write-host "View $logFilePath for summary of results. `r`n Import $csvFilePath into KeePass to share usernames and passwords with Personnel. `r`n Please delete the CSV file immediately after import."
}

remove-pssession $session