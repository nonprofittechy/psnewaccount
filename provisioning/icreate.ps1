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
		if($userinfo['department'] -eq "CASLS") {
			$userinfo.officeID = 2
			$userinfo.generalNo="(617) 603-2700"
			$userinfo.faxNo = "(617) 494-8222"
		} else {
			$userinfo.officeID = 1
			$userinfo.generalNo = "(617) 371-1234"
			$userinfo.faxNo = "(617) 371-1222"
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