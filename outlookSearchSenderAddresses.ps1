Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$namespace = new-object -comobject outlook.application
$MAPI = $namespace.GetNamespace("MAPI")
$Inbox = $MAPI.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Items
# iterate mail items, returning sender
$Senders = `
    foreach ( $MailItem in $Inbox ) { 
        $MailItem.SenderEmailAddress 
    }
$namespace.Quit()
$SendersFiltered = `
($Senders | select -Unique) | where { 
 $pref = $_.split("@")[0]; 
 $suff = $_.split("@")[1]; 
 $boolWordMatch = $False;
 # per-user custom dictionary 
 $arrStrDict = ( `
 'support','ship','info','service','billing','customer','account','sales', `
 'reservation','email','upgrade','message','subscribe'
 )
 foreach ( $dictionaryWordEx in $arrStrDict ) {
  if ($pref -match $dictionaryWordEx ) { $boolWordMatch = $True 
  } 
 }
 # general dictionary of computerized address indicators
 $boolCharMatch = $False
 $arrCharDict = ( "-", "=", "+" )
 foreach ( $dictionaryChar in $arrCharDict ) {
  if ( $pref.contains($dictionaryChar) ) { $boolCharMatch = $True 
  }
 }
 $boolWordMatch -eq $False -and `
 $boolCharMatch -eq $False -and `
 $suff -ne $null -and `
 $pref -notmatch $suff.split(".")[-2]
}
# sample output
$SendersFiltered | select -first 1
