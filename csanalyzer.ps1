<#
The sample scripts are not supported under any Microsoft standard support program or 
service. The sample scripts are provided AS IS without warranty of any kind. Microsoft 
further disclaims all implied warranties including, without limitation, any implied 
warranties of merchantability #or of fitness for a particular purpose. The entire risk 
arising out of the use or performance of #the sample scripts and documentation remains 
with you. In no event shall Microsoft, its authors, or anyone else involved in the 
creation, production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability
to use the sample scripts or documentation, even if Microsoft has been advised of the 
possibility of such damages.
#>

Param(
	[Parameter(Mandatory=$true, HelpMessage="Must be a file generated using csexport 'Name of Connector' export.xml /f:x)")]
	[string]$xmltoimport,
	[Parameter(Mandatory=$false, HelpMessage="Maximum number of users per output file")][int]$batchsize=1000,
	[Parameter(Mandatory=$false, HelpMessage="Show console output")][bool]$showOutput=$false
)


#LINQ isn't loaded automatically, so force it
[Reflection.Assembly]::Load("System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089") | Out-Null

#internal variables
[int]$count=1
[int]$outputfilecount=1
[array]$objOutputUsers=@()
[string]$resolvedXMLtoimport
[int]$countadds=0
[int]$countdeletes=0
[int]$countupdates=0

#XML must be generated using "csexport "Name of Connector" export.xml /f:x"
write-host "Importing XML" -ForegroundColor Yellow

#XmlReader.Create won't properly resolve the file location,
#so expand and then resolve it
$resolvedXMLtoimport=Resolve-Path -Path ([Environment]::ExpandEnvironmentVariables($xmltoimport))

#use an XmlReader to deal with even large files
$result=$reader = [System.Xml.XmlReader]::Create($resolvedXMLtoimport) 
$result=$reader.ReadToDescendant('cs-object')

#start processing the XML
do 
{	
	#create the object placeholder
	#adding them up here means we can enforce consistency
	$objOutputUser=New-Object psobject
	Add-Member -InputObject $objOutputUser -MemberType NoteProperty -Name ID -Value ""
	Add-Member -InputObject $objOutputUser -MemberType NoteProperty -Name Type -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name DN -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name operation -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name UPN -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name displayName -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name sourceAnchor -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name alias -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name primarySMTP -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name onPremisesSamAccountName -Value ""
	Add-Member -inputobject $objOutputUser -MemberType NoteProperty -Name mail -Value ""
	
	$user = [System.Xml.Linq.XElement]::ReadFrom($reader)
	if ($showOutput) {Write-Host "Found an exported object..." -ForegroundColor Green}
	
	#object id
	$outID=$user.Attribute('id').Value
	if ($showOutput) {Write-Host "ID: ${outID}"}
	$objOutputUser.ID=$outID
	
	#object type
	$outType=$user.Attribute('object-type').Value
	if ($showOutput) {Write-Host "Type: ${outType}"}
	$objOutputUser.Type=$outType
	
	#dn
	$outDN= $user.Element('unapplied-export').Element('delta').Attribute('dn').Value
	if ($showOutput) {Write-Host "DN: ${outDN}"}
	$objOutputUser.DN=$outDN
	
	#operation
	$outOperation= $user.Element('unapplied-export').Element('delta').Attribute('operation').Value
	if ($showOutput) {Write-Host "Operation: ${outOperation}"}
	$objOutputUser.operation=$outOperation

	#track the operation count
	switch ($outOperation) 
	{
		"update"
		{
			$countupdates++
		}
		"delete"
		{
			$countdeletes++
		}
		"add"
		{
			$countadds++
		}
	}
	
	#now that we have the basics, go get the details

	foreach ($attr in $user.Element('unapplied-export-hologram').Element('entry').Elements("attr"))
	{
		$attrvalue=$attr.Attribute('name').Value
		$internalvalue= $attr.Element('value').Value
	
		switch ($attrvalue)
		{
			"userPrincipalName"
			{	
				if ($showOutput) {Write-Host "UPN: ${internalvalue}"}
				$objOutputUser.UPN=$internalvalue
			}
			"displayName"
			{
				if ($showOutput) {Write-Host "displayName: ${internalvalue}"}
				$objOutputUser.displayName=$internalvalue
			}
			"sourceAnchor"
			{
				if ($showOutput) {Write-Host "sourceAnchor: ${internalvalue}"}
				$objOutputUser.sourceAnchor=$internalvalue
			}
			"alias"
			{
				if ($showOutput) {Write-Host "alias: ${internalvalue}"}
				$objOutputUser.alias=$internalvalue
			}
			"proxyAddresses"
			{
				if ($showOutput) {Write-Host "primarySMTP: (${internalvalue} -replace "SMTP:","")"}
				$objOutputUser.primarySMTP=$internalvalue -replace "SMTP:",""
			}
		}
	}
	
	$objOutputUsers += $objOutputUser
	
	Write-Progress -activity "Processing ${xmltoimport} in batches of ${batchsize}" -status "Batch ${outputfilecount}: " -percentComplete (($objOutputUsers.Count / $batchsize) * 100)

	
	#every so often, dump the processed users in case we blow up somewhere
	if ($count % $batchsize -eq 0)
	{
		Write-Host "Hit the maximum users processed in a batch..." -ForegroundColor Yellow

		#export the collection of users as as CSV
		Write-Host "Writing processedusers${outputfilecount}.csv" -ForegroundColor Yellow
		$objOutputUsers | Export-Csv -path .\processedusers${outputfilecount}.csv -NoTypeInformation 

		#increment the output file counter
		$outputfilecount++

		#reset the collection and the user counter
		$objOutputUsers = $null
		$count=0
	}
	
	$count++	
	
	#need to bail out of the loop if no more users to process
	if ($reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement)
	{
		break
	}
	
	
	
} while ($reader.Read)

#need to write out any users that didn't get picked up in a batch of 1000
#export the collection of users as as CSV
Write-Host Writing processedusers${outputfilecount}.csv -ForegroundColor Yellow
$objOutputUsers | Export-Csv -path processedusers${outputfilecount}.csv -NoTypeInformation 

#report the summary data
Write-Host "Summary:"
Write-Host "`tFound ${countadds} add operations"
Write-Host "`tFound ${countdeletes} delete operations"
Write-Host "`tFound ${countupdates} update operations"
