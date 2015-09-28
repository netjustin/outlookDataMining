Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$namespace = new-object -comobject outlook.application
$MAPI = $namespace.GetNamespace( "MAPI" )
$Inbox = $MAPI.GetDefaultFolder( [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox ).Items
# iterate mail items, returning sender
$objSenders = `
 foreach ( $MailItem in $Inbox ) { 
  $MailItem.SenderEmailAddress 
 }
$namespace.Quit()
$objSendersFiltered = `
( $objSenders | select -Unique ) | where { 
 $strPrefix = $_.split("@")[0]; 
 $strSuffix = $_.split("@")[1]; 
 $boolWordMatch = $False;
 # per-user custom dictionary 
 $arrStrFilterWords = ( `
 'support','ship','info','service','billing','customer','account', `
 'sales','reservation','email','upgrade','message','subscribe'
 )
 foreach ( $dictionaryWordEx in $arrStrFilterWords ) {
  if ( $strPrefix -match $dictionaryWordEx ) { 
   $boolWordMatch = $True 
   break
  } #if
 } #foreach
 # char indicators of computerized address 
 $boolCharMatch = $False
 $arrCharDict = ( "-", "=", "+" )
if ( $boolWordMatch -ne $True ) {
 foreach ( $dictionaryChar in $arrCharDict ) {
  if ( $strPrefix.contains( $dictionaryChar ) ) { 
   $boolCharMatch = $True 
   break
   } #if
  } #foreach
 } #if
 $boolWordMatch -eq $False -and `
 $boolCharMatch -eq $False -and `
 $strSuffix -ne $null -and `
 $strPrefix -notmatch $strSuffix.split(".")[-2]
}
# sample output
$objSendersFiltered | select -first 1
