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
  Login username for Legal Server with "add/edit users" permission
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
import-csv .\users.csv | new-legalserveraccount -username [USERNAME] -password [PASSWORD] -loginURL "https://contoso.legalserver.org" -newUserURL "https://contoso.legalserver.org/user/process/dynamic_XX"
.EXAMPLE
new-legalserveraccount -adminUsername [USERNAME] -adminPassword [PASSWORD] -loginURL "https://contoso.legalserver.org" -newUserURL "https://contoso.legalserver.org/user/process/dynamic_XX"
.NOTES
	Author: Quinten Steenhuis, 12/22/2017
  Reference: https://cmdrkeene.com/automating-internet-explorer-with-powershell
  see also: https://www.sepago.com/blog/2016/05/03/powershell-exception-0x800a01b6-while-using-getelementsbytagname-getelementsbyname
  https://msdn.microsoft.com/en-us/library/aa752084(v=vs.85).aspx
  
  Licensed under GPLv3. See included license.md in this folder or https://www.gnu.org/licenses/gpl.html
.LINK
https://github.com/nonprofittechy/psnewaccount
#>
[CmdletBinding()]
Param(
  [Parameter(Mandatory = $True,HelpMessage="HomePage URL")][string]$loginURL,
  [Parameter(Mandatory = $True,HelpMessage="New User Process URL")][string]$newUserURL,
  [Parameter(Mandatory = $True)][string]$adminUsername,
  [Parameter(Mandatory = $True)][string]$adminPassword,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$username,
  [Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName=$True)][string]$password,
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
  
  # Legal Server displays a list with the class name "form_errors" if anything goes wrong
  function get-legalservererrors ($ie) {
    return ($ie.document.IHTMLDocument3_getElementsByTagName("ul")) | %{if ($_.className -eq "form_errors") {$_.innerHtml} }
  }
  
  function test-legalservererrors ($ie) {
    $e = get-legalservererrors $ie
    if ( $e ) {
      throw ($e -replace '<[^>]+>|(\&nbsp\;)','') + " : Skipping Legal Server account for username " +  $username + ", Name: " + $givenname + " " + $surname # clean the HTML tags from the error in Legal Server
    }
  }
  
  ########################################################
  # Customize the variable below if needed for your Legal Server site. Many of the field 
  # names appear to be standardized and may not need to be altered.
  # double-check the "Do Not Set Organization" button ID for your site.
  $fields = 
  @{
    "login_form" = "login_form";
    "login_username" = "____login_login";
    "login_password" = "____login_password";
    "login_submit" = "____login_submit";
    "user_login" = "user:login";
    "user_type" = "user:types";
    "user_role" = "user:role";
    "user_office" = "user:office_id";
    "user_program" = "user:program_id";
    "user_expiration_date" = "user:date_end";
    "user_first_name" = "contact:first";
    "user_middle_name" = "contact:middle";
    "user_last_name" = "contact:last";
    "user_email" = "user:email";
    "user_phone" = "user:phone_business";
    "organization_donotsetbutton" = "form_element_toggler_5_input";
    "user_password" = "____password";
    "user_confirm_password" = "____confirm_password";
  }

  # Create new IE COM object and login to Legal Server
  $ie = New-Object -com internetexplorer.application;
  if ($PSBoundParameters.debug) {
    $ie.visible = $true;
  }
  $ie.navigate($loginUrl);
  while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
  
  # check for the login form - if it's not present, check to see if there is an "admin" link. If not--something went wrong.
  if ( ($ie.Document.IHTMLDocument3_getElementsByName($fields['login_form']) | select -first 1) ) {
    ($ie.document.IHTMLDocument3_getElementsByName($fields['login_username']) |select -first 1).value = $adminUsername
    ($ie.document.IHTMLDocument3_getElementsByName($fields['login_password']) |select -first 1).value = $adminPassword
    ($ie.document.IHTMLDocument3_getElementsByName($fields['login_submit']) |select -first 1).click()
  } elseif ( $ie.Document.IHTMLDocument3_getElementsByTagName("a") | where {$_.innerText -eq "Admin"} | select -first 1) {
    # we are already logged in
  } else {
    throw "Something went wrong trying to access the main login page."
  }
  
  while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
  #Start-sleep -seconds 10
}

Process {
  try {

  $ie.navigate($newUserURL);
  while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
  
  # Page 1 - System Information
  ($ie.document.IHTMLDocument3_getElementsByName($fields['user_login']) |select -first 1).value = ($username.tolower())
  if ($accountExpirationDate) {
    $formattedDate = get-date $accountExpirationDate -format "MM/dd/yyyy"
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_expiration_date']) |select -first 1).value = $formattedDate
  }
  ($ie.document.IHTMLDocument3_getElementsByName($fields['user_type']) |select -first 1).value = $LSTypeID
  ($ie.document.IHTMLDocument3_getElementsByName($fields['user_role']) |select -first 1).value = $LSRoleID
  ($ie.document.IHTMLDocument3_getElementsByName($fields['user_office']) |select -first 1).value = $LSOfficeID
  ($ie.document.IHTMLDocument3_getElementsByName($fields['user_program']) |select -first 1).value = $LSProgramID
  
  write-debug "Submitting Page 1"
  
  ($ie.document.IHTMLDocument3_getElementsByName("submit_button") |select -first 1).click()
  while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
  
  test-legalservererrors $ie
   
  #Page 2 - Contact Information
  # Check to see if we have a first name field
  if ( ($ie.document.IHTMLDocument3_getElementsByName($fields['user_first_name']) |select -first 1) ) {
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_first_name']) | select -first 1).value = $givenName
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_last_name']) | select -first 1).value = $surname
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_middle_name']) | select -first 1).value = $middleName
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_email']) | select -first 1).value = $emailAddress
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_phone']) | select -first 1).value = $phone
    
    write-debug "Submitting Page 2"  
    ($ie.document.IHTMLDocument3_getElementsByName("submit_button") |select -first 1).click()
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
    
    test-legalservererrors $ie
  
    
    
  } else {
    write-error "Unexpected page - expected to see a field named 'contact:first'."
    break;
  }
  
  #Page 3 - Organization Affiliation
  if (($ie.document.IHTMLDocument3_getElementById($fields["organization_donotsetbutton"])) ) {
    # click the "Do not set Organization" button
    ($ie.document.IHTMLDocument3_getElementById($fields["organization_donotsetbutton"])).click()
   
    write-debug "Submitting Page 3"
    ($ie.document.IHTMLDocument3_getElementsByName("submit_button") |select -first 1).click()
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
    
    test-legalservererrors $ie
   
  } else {
      write-error "Unexpected page - expected to be on the Organization Affiliation page"
      break;
  }
  
  #Page 4 - Password
  if (($ie.document.IHTMLDocument3_getElementsByName($fields['user_password']))) {
    # Set the "Needs to change password" radio to Yes
    ($ie.document.IHTMLDocument3_getElementsByTagName("input") | where {$_.type -eq "radio" -and $_.value -like "t"} | select -first 1).checked = $true
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_password']) | select -first 1).value = $password
    ($ie.document.IHTMLDocument3_getElementsByName($fields['user_confirm_password']) | select -first 1).value = $password
    
    write-debug "Submitting Page 4"
    ($ie.document.IHTMLDocument3_getElementsByName("submit_button") |select -first 1).click()
    
    while ($ie.Busy -eq $true) { Start-Sleep -Seconds 1; }
    
    test-legalservererrors $ie

  } else {
    write-error "Unexpected page - expected to be on the Password page."
    break;
  }
  
  } catch {
    write-error $_
  }
  
}

End {
  write-debug "Closing IE"
  $ie.quit()
}