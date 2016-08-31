############################################
# Office 365 functions

# creates a new mailbox with the tenant address set in provision.variables.ps1
# See: http://www.office365tipoftheday.com/2014/01/31/remote-mailboxes/
function new-office365mailbox ($userinfo, $domaincontroller) {
	$routingAddress = $userinfo.samAccountName + $Office365TenantAddressSuffix
	enable-remotemailbox ($userinfo.upn) -primarysmtpaddress ($userinfo.emailAddress) -remoteroutingaddress $routingaddress -domaincontroller $domaincontroller
	set-remotemailbox -identity ($userinfo.upn) -emailAddresses @{add=$routingAddress} -domaincontroller $domaincontroller
}


function test-mailboxenabled($samAccountName) {
	$u = get-aduser $samaccountname -properties Mail
	
	return ! ([bool]($u.Mail -eq $null))
}
