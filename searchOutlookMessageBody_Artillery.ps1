<# 	Enumerate an inbox full of known Artillery alerts, and parse out the 
	relevant body data.
	Known limitation: Does not parse the SMTP header, which would provide a 
	reasonable guess of what addr_dst was.
 #>

 <# global(s)	#>

$outFile = "Artillery.csv"


<# helper func	#>

function ArtilleryEmailBodyToCsv($s) 
{	# trim and decorate elements of the alert
	$s = $s -replace '\[!\] Artillery has detected an attack from IP address: ',''
	$s = $s -replace 'for a connection on a honeypot port: ',''
	$s = $s -replace ' ','","'
	$s = "{0}{1}{0}" -f [char]34, $s	# enquote
	return $s
}


<# main			#>

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$namespace 	= new-object -comobject outlook.application
$MAPI 		= $namespace.GetNamespace("MAPI")
$Inbox 		= $MAPI.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Items
# iterate mail items, returning header
$Bodies 	= `
	foreach ( $MailItem in $Inbox ) { 
		if 	( 	$MailItem 		-ne $Null -and `
				$MailItem.Body 	-ne $Null) { 
					$MailItem.Body 	# return non-null items
				}
	}
$namespace.Quit()


<# filtering	#>

$CsvBodies = `
	foreach ( $Body in $Bodies ) { 
		ArtilleryEmailBodyToCsv( $Body )
	}
	
<# output		#>

'"date","time","addr_src","port_dst"' 	| Out-File -Encoding ASCII -FilePath $outFile
$CsvBodies 								| Out-File -Append -Encoding ASCII -FilePath $outFile
