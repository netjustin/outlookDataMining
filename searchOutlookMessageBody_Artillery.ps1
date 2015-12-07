<# 	Enumerate an inbox full of known Artillery alerts, and parse out the 
	relevant body data.
	Known limitation: Does not parse the SMTP header, which would provide a 
	reasonable guess of what addr_dst was.
 #>

 <# global(s)	#>

$outFile 			= "Artillery.csv"
# $DeleteAfterExport 	= $True
$ConstrainResults	= $False
$ConstrainCount		= 10

<# helper func	#>

function ArtilleryEmailBodyToCsv($s) 
{	# trim and decorate socket pair elements from body
	$r = '\[!\] Artillery has (detected an attack from|blocked \(and blacklisted\) the) IP address: '
	$s = $s -replace $r, ''
	$r = 'for (a connection on|connecting to) a honeypot (restricted )?port: '
	$s = $s -replace $r,''
	$s = $s -replace ' ','","'
	$s = "{0}{1}{0}" -f '"', $s	# enquote
	return $s
}


<# main			#>

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$namespace 	= new-object -comobject outlook.application
$MAPI 		= $namespace.GetNamespace("MAPI")
$Inbox 		= $MAPI.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Items
# iterate mail items, returning header
$i=0
$Bodies 	= `
	foreach ( $MailItem in $Inbox ) { 
		if 	( 	($i++ -lt $ConstrainCount -or $ConstrainResults -eq $False) -and `
				$MailItem 		-ne $Null -and `
				$MailItem.Body 	-ne $Null -and `
				$MailItem.Subject -match 'Artillery has detected an attack' )
				{ 	# return matching items
					$MailItem.Body 	
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
