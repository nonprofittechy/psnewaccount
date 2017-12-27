##################################################################
#
#				Account Creation Sub-Functions
#
##################################################################

Function OutputLog($textOrObject) {
	$textOrObject >> $global:logFilePath
}

function isValidDepartment($department) {
	return $depts.contains($department)
}

function isValidManager($manager) {
	return $manager -ne ""
}

function getManagerSAMAccountName($manager, $adDomainControllerHostName){
	if ($manager.trimEnd() -eq "") {
		OutputLog("*** WARNING *** : Manager name not provided.")
		return ""
	}
	if ((test-user -samaccountname $manager -addomaincontrollerhostname $adDomainControllerHostName)) {
		return $manager
	} else {
		# Let's see if the name was mistakenly given in the format "First Last" or "Last First"
		$names = $manager.split("\s+")
		if ($names -is [system.array]) {
			$firstlast = get-aduser -server $adDomainControllerHostName -filter {sn -like $names[1] -and givenname -like $names[0]}
		
			if ($firstlast -ne $null -and !($firstlast -is [array]) ) {
				return $firstlast.samAccountName
			}
			
			$lastfirst = get-aduser -filter {sn -like $names[0] -and givenname -like $names[1]}
			if ($lastfirst -ne $null -and !($lastfirst -is [array]) ) {
				return $lastfirst.samAccountName
			}
		}
		OutputLog("*** Warning: Manager name " + $manager + " appears to be ambiguous or invalid. Please provide a user id, such as jbowman. Will try using the template's manager instead.")
		return ""
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


####################################################
# 	Verify account creation
# 	output to the log file and add to the array for CSV to import into KeePass / sending of emails
function VerifyAccount($user, $adDomainControllerHostName) {
	if ($adDomainControllerHostName -eq $null) {
		$adDomainControllerHostName = get-addomaincontroller | select -expand "HostName"
	}
	if (!(test-user -samAccountName $user.samAccountName -adDomainControllerHostName $adDomainControllerHostName )) {
		OutputLog("*** WARNING *** User's Active Directory account "  + $user.name + " (" + $user.samAccountName + ")" + " appears not to have been created. Please wait a few minutes and verify account creation. Mailbox may need to be created manually.")
	} else {
		OutputLog("Successfully created new user in Active Directory: " + $user.name + ", login name: " + $user.samAccountName + ", Department: " + $user.department )
    
		if ($user.accountExpirationDate) {
      $formattedDate = get-date $user.accountExpirationDate -format "MM/dd/yyyy"
    }
    
		$userObj = @{"Title" = $user.name; 
				   "Username" = $user.samAccountName;
				   "Password" = $user.clearTextPassphrase;
				   "Url" = "";
				   "Notes" = ("Unit: " + $user.Unit + "`r`n Shared Folder: " + $user.userSharedFolder);
				   "samAccountName" = $user.samAccountName;
				   "Department" = $user.department;
				   "emailAddress" = $user.emailAddress;
				   "role" = $user.title;
				   "accountExpirationDate" = $formattedDate;
				   "SharedFolder" = $user.userSharedFolderOther;
           "GivenName" = $user.givenName;
           "Surname" = $user.sn;
           "origUnit" = $user.origUnit;
           "middleName" = $user.mi;
           "phone" = $user.phone
				  }
	
		# $global:userSummary += new-object psobject -property $userObj
		return new-object psobject -property $userObj
			
	}
	
	if (test-path function:test-mailboxenabled) {
		if (!(test-mailboxenabled($user.samAccountName))) {
			OutputLog("*** WARNING *** User account " + $user.name + "(" + $user.samAccountName + ")" + " couldn't be found. The account may not be mail-enabled, or there may just be a lag in syncing between domain controllers.")
		} else {
			OutputLog("Successfully requested an Exchange Mailbox for "  + $user.name + " (" + $user.samAccountName + ") Mailbox should be created at the next synchronization with Office 365 (every 3 hours).")
		}
	} 
}

##############################
# Generate a unique account name given user info
# Try FLastname, followed by FMLastname, followed by FirstLastname, followed by FirstMLastname, followed by FL#
# strip out any non alpha-numeric characters along the way and any leading digits
Function UniqueSamAccountName($userinfo, $adDomainControllerHostName) {	

	#write-host ($userinfo | out-string)

	if ($adDomainControllerHostName -eq $null) {
		$adDomainControllerHostName = get-addomaincontroller | select -expand "HostName"
	}
	
	$gn = ($userinfo.givenName -replace  '[^a-zA-Z0-9]','') -replace '^\d+'
	$sn = ($userinfo.sn -replace '[^a-zA-Z0-9]','') -replace '^\d+'
	$mi = ($userinfo.mi -replace  '[^a-zA-Z0-9]','') -replace '^\d+'
	
	$samAccountName = ($userinfo.givenName.substring(0,1) + $sn)
	if (test-user -adDomainControllerHostName $adDomainControllerHostName -samAccountName $samAccountName) {
		$samAccountName = ($gn.substring(0,1) + $mi + $sn)
		if ( test-user -adDomainControllerHostName $adDomainControllerHostName -samAccountName $samAccountName ){
			$samAccountName = ($gn + $sn)
			if ( test-user -adDomainControllerHostName $adDomainControllerHostName -samAccountName $samAccountName ){
				$samAccountName = ($gn + $mi + $sn)	 		
				$i = 2 #  Logically it's at least the second user with this name--avoid confusion, start numbering at 2!
				while (test-user -adDomainControllerHostName $adDomainControllerHostName -samAccountName $samAccountName) {
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
function UniqueName($userinfo, $adDomainControllerHostName) {

	if ($adDomainControllerHostName -eq $null) {
		$adDomainControllerHostName = get-addomaincontroller | select -expand "HostName"
	}
	
	$dc = $adDomainControllerHostName

	$Name = ($userinfo.sn + ", " + $userinfo.givenName)
	
	if (Get-ADUser -filter {name -like $name} -server $dc) {
		if ($userinfo.mi -ne "") {
			$name += " " + $userinfo.mi
		}
		$i = 2
		if (get-ADUser -filter {name -like $name}  -server $dc) {
			$name = ($userinfo.sn + ", " + $userinfo.givenName + " (" + $userinfo.samAccountName + ")")
		}
		while (Get-ADUser -filter {name -like $name} -server $dc) {
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

# Export the user summary we've been building to a CSV file
Function export-keepassCSV ($path) {
	$global:userSummary | select-object Username,Title,Url,Password,Notes | export-csv -path $path -noTypeInformation
}


