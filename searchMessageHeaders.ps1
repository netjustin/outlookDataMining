$MatchString = "X-Mailer: YahooMailWebService/0.8.201.700"
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$namespace = new-object -comobject outlook.application
$MAPI = $namespace.GetNamespace("MAPI")
$Inbox = $MAPI.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Items
# iterate mail items, returning header
$Headers = `
    foreach ( $MailItem in $Inbox ) { 
        $MailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") 
    }
$namespace.Quit()   # at this point, variable $Headers contains all Inbox headers
$MatchingHeaders = $Headers | where { $_.contains( $MatchString ) }
$MatchingHeaders | Select-Object -First 1   # sample output
