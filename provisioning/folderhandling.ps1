#################################################
# Get a path name for the user's shared folder to pass to the folder creation function
# based on user's department, SITE1/City1, and intern/full time staff
function getUserSharedFolderPath($ht) {

		if ($ht.department -eq "SITE1") {
			$userSharedFolder = $City1Share
		} else {
			$userSharedFolder = $City2Share
		}
		
		if (($ht.title -match "Intern" ) -or ($ht.title -match "Student")) {
			$userSharedFolder += $depts[$ht.department].students + $ht.samAccountName
		} else {
			$userSharedFolder += $depts[$ht.department].folder + $ht.samAccountName
		}
		
		return [string]$userSharedFolder
}

function getUserSharedFolderPathOther($ht) {

		if ($ht.department -eq "SITE1") {
			$userSharedFolderOther = $City1Drive
		} else {
			$userSharedFolderOther = $City2Drive
		}
		
		if (($ht.title -match "Intern" ) -or ($ht.title -match "Student")) {
			$userSharedFolderOther += $depts[$ht.department].students + $ht.samAccountName
		} else {
			$userSharedFolderOther += $depts[$ht.department].folder + $ht.samAccountName
		}
		
		return [string]$userSharedFolderOther
}

#########################
# Create home folder / unit subfolder
function CreateHomeFolder($userinfo) {
	$path = ""
	
	# Create the Unit folder for the user (on S: drive)
	if ($userinfo["department"] -eq "SITE1") {
		$path = $City1Share
	} else {
		$path = $City2Share
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