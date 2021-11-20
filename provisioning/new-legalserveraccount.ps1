<#
.SYNOPSIS
  Creates a new account in Legal Server using IE automation.
.DESCRIPTION
This should be used carefully as users cannot be deleted from Legal Server. The list of field names may need to be customized for your Legal Server site. Check the variable $fields in this script for field names that may need to be updated. This script includes basic error handling and will skip invalid user information. However, it will not resolve errors, such as usernames being non-unique.
.PARAMETER loginURL
  The URL for the login screen in Legal Server
.PARAMETER newUserURL
  The URL for the first screen of the new user process
.PARAMETER adminUsername
  Login username for Legal Server with "both internal and external API add/edit users" permission
.PARAMETER adminPassword
  Password matching username
.PARAMETER username
  Name of the account to create in Legal Server.
  Can be provided on pipeline.
.PARAMETER password
  Password for new account.
  Can be provided on pipeline.
.PARAMETER givenName
  Can be provided on pipeline.
.PARAMETER surname
  Can be provided on pipeline.
.PARAMETER middleName
  Can be provided on pipeline.
.PARAMETER LSProgramID
  Program for new account (String), must match internal database value (visible in HTML SELECT) 
  in Legal Server for corresponding Program lookup value.
  Can be provided on pipeline.
.PARAMETER LSTypeID
  Internal database value for Legal Server user "Type"
.PARAMETER LSRoleID
  Internal database value for Legal Server user "Role"
.PARAMETER LSOfficeID
  Internal database value for Legal Server user "Office"
.EXAMPLE
Example with list of user fields in a .CSV file for piped input:
import-csv .\users.csv | new-legalserveraccount -username [USERNAME] -password [PASSWORD] -loginURL "https://contoso.legalserver.org"
.EXAMPLE
new-legalserveraccount -adminUsername [USERNAME] -adminPassword [PASSWORD] -loginURL "https://contoso.legalserver.org" 
.NOTES
	Author: Quinten Steenhuis, 12/22/2017
  Migrated to new Legal Server API on 11/20/2021
  see also: https://www.sepago.com/blog/2016/05/03/powershell-exception-0x800a01b6-while-using-getelementsbytagname-getelementsbyname
  https://msdn.microsoft.com/en-us/library/aa752084(v=vs.85).aspx
  
  Licensed under GPLv3. See included license.md in this folder or https://www.gnu.org/licenses/gpl.html
.LINK
https://github.com/nonprofittechy/psnewaccount
#>
[CmdletBinding()]
Param(
  [Parameter(Mandatory = $True,HelpMessage="Legal Server base URL, like: https://myorg-demo.legalserver.org")][string]$loginURL,
  [Parameter][string]$newUserURL, # Not used, left in for possible backwards compatibility reasons
  [Parameter(Mandatory = $True)][string]$adminUsername,
  [Parameter(Mandatory = $True)][string]$adminPassword,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$username,
  [Parameter(ValueFromPipelinebyPropertyName=$True)][string]$password, # Not used
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$LSProgramID,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$LSTypeID,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$LSRoleID,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$LSOfficeID,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$givenName,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$surname,
  [Parameter(ValueFromPipelinebyPropertyName=$True)][string]$middleName,
  [Parameter(ValueFromPipelinebyPropertyName=$True)][string]$phone,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$emailaddress,
  [Parameter(ValueFromPipelinebyPropertyName=$True)][string]$accountExpirationDate
  
)
Begin {

}

Process {
  try {

  $user_params = @{ 
    "first"= $givenName
    "middle"= $middleName
    "last"= $surname
    "login"= $username
    "email"= $emailaddress
    "role"= $LSRoleID
    "types"= $LSTypeID
    "office"= $LSOfficeID
    "program"= $LSProgramID
    "password_reset"= $true 
  }    

  invoke-webrequest -Method POST -Uri ($loginURL + "/api/v1/users") -Headers @{ Authorization = "Basic "+ [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($adminUsername):$($adminPassword)")); "Content-Type"="application/json"} -Body ($user_params | convertto-json)
  
  } catch {
    write-error $_
  }
  
}

End {
  write-debug "All done"
}
