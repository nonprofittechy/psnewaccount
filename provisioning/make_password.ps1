#Given a dictionary, generate a random 2-word passphrase. The first letter is capitalized and a random number is appended to meet
# complexity requirements
function PassPhrase ($words) {
	
	$phrase = (get-random $words) + " " + (get-random $words)
	
	$phrase = ( ($phrase.subString(0,1).toUpper()) + $phrase.substring(1) + (get-random -max 99))
	
	return [string]$phrase
}

$dictionaryPath =  (split-path $myinvocation.mycommand.path) + "\dict.txt"
if (test-path $dictionaryPath) {
	$dictionary = get-content $dictionaryPath
} else {
	throw "Dictionary file for passphrase generation '$dictionaryPath' could not be found."
}

for ($i = 0; $i -lt 10; $i++) {
	write-host (PassPhrase($dictionary))
}