Get-ChildItem -Path ".\functions" -Filter *.ps1 | ForEach-Object {
    . $_.FullName
}

$channelObjList = @()
$fileChannels = @()
$jdbcChannels = @()
$restChannels = @()
$soapChannels = @()
$sftpChannels = @()
$idocChannels = @()
$otherChannels = @()

$properties = Get-Content .\Properties.json | ConvertFrom-Json

$url = $properties.pi_system_url
if ( $null -ne $properties.file_name) {
    $outputpath = $properties.file_name
}
else {
    $outputpath = $outputpath = join-path -Path $env:USERPROFILE -ChildPath "desktop\PIPOChannelDetails.xlsx"
}
$download_channel_xml = $properties.download_channel_xml

[Xml] $SOAPRequest = Get-Content .\ChannelQueryRequest.xml
[Xml] $SOAPReadChannelRequest = Get-Content .\ChannelReadRequest.xml

# Request for credentials to autheticate the PI channel service
$credential = Get-Credential
$bytes = [System.Text.Encoding]::UTF8.GetBytes(
    ('{0}:{1}' -f $credential.UserName, $credential.GetNetworkCredential().Password)
)
$authorization = 'Basic {0}' -f ([Convert]::ToBase64String($bytes))

Write-Host 'Reading Channel Names.....'
# Call the web-service to read all the channels
$ReturnXml = CallPIChannelService $SOAPRequest $authorization $url

foreach ($channel in $ReturnXml.GetElementsByTagName('rn3:CommunicationChannelID')) {
    $channelObj = New-Object -TypeName PSObject
    # Write-Output $channel.ChannelID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'PartyID' -Value $channel.PartyID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ComponentID' -Value $channel.ComponentID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ChannelID' -Value $channel.ChannelID
    $channelObjList += $channelObj
}

$channelXmlFilePath = '.\Channels' 
Write-Host 'Reading Channel XML configuration.....'
foreach ( $channel in $channelObjList ) {
    $request = [Xml] $SOAPReadChannelRequest.OuterXml.Replace("{{PartyID}}", $channel.PartyID).Replace("{{ComponentID}}", $channel.ComponentID).Replace("{{ChannelID}}", $channel.ChannelID)

    $ReturnXml = CallPIChannelService $request $authorization $url
    if($download_channel_xml){
        if($channel.PartyID){
            $fileName = "{0}\{1}_{2}_{3}.{4}" -f $channelXmlFilePath, $channel.PartyID, $channel.ComponentID, $channel.ChannelID, 'xml'
        } else {
            $fileName = "{0}\{1}_{2}.{3}" -f $channelXmlFilePath, $channel.ComponentID, $channel.ChannelID, 'xml'
        }
        $ReturnXml.OuterXml | Out-File $fileName
    }

    $direction = $ReturnXml.GetElementsByTagName('rn3:CommunicationChannel').Direction
    $adapterEngine = $ReturnXml.GetElementsByTagName('rn3:CommunicationChannel').AdapterEngineName
    $adapterType = $ReturnXml.GetElementsByTagName('rn3:AdapterMetadata').name
    $channelObj = New-Object -TypeName PSObject
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'PartyID' -Value $channel.PartyID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ComponentID' -Value $channel.ComponentID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ChannelID' -Value $channel.ChannelID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'Direction' -Value $direction
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'AdapterType' -Value $adapterType
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'AdapterEngine' -Value $adapterEngine

    switch ($adapterType) {
        'File' {
            $channelObj = ProcessFileAdapterParameters $ReturnXml.GetElementsByTagName('rn3:AdapterSpecificAttribute') $channelObj
            $fileChannels += $channelObj
        }
        'JDBC' {
            $channelObj = ProcessJdbcAdapterParameters $ReturnXml.GetElementsByTagName('rn3:AdapterSpecificAttribute') $channelObj
            $jdbcChannels += $channelObj
        }
        'REST' {
            $channelObj = ProcessRestAdapterParameters $ReturnXml.GetElementsByTagName('rn3:AdapterSpecificAttribute') $channelObj
            $restChannels += $channelObj
        }
        'SFTP' {
            $channelObj = ProcessSftpAdapterParameters $ReturnXml.GetElementsByTagName('rn3:AdapterSpecificAttribute') $channelObj
            $sftpChannels += $channelObj
        }
        'SOAP' {
            $channelObj = ProcessSoapAdapterParameters $ReturnXml.GetElementsByTagName('rn3:AdapterSpecificAttribute') $channelObj
            $soapChannels += $channelObj
        }
        'IDoc_AAE' {
            $channelObj = ProcessIdocAdapterParameters $ReturnXml.GetElementsByTagName('rn3:AdapterSpecificAttribute') $channelObj
            $idocChannels += $channelObj
        }
        Default {
            $channelObj = ProcessOtherAdapterTypesParameters $ReturnXml.GetElementsByTagName('rn3:AdapterSpecificAttribute') $channelObj
            $otherChannels += $channelObj
        }
    }
}

Write-Host 'Preparing Excel Output...'

#open excell
$excel = New-Object -ComObject excel.application
$excel.visible = $false
$workbook = $excel.Workbooks.Add()

if ($otherChannels) {
    #give the remaining worksheet a name
    $otherChannelsSheet = $workbook.Worksheets.Item(1)
    $otherChannelsSheet.name = 'Other Channels'
    $otherChannelsSheet = OtherAdapterChannelsToExcel $otherChannels $otherChannelsSheet
    #adjusting the column width so all data's properly visible
    $otherChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
}

if ($idocChannels) {
    $idocChannelsSheet = $workbook.Worksheets.add()
    $idocChannelsSheet.name = 'Idoc Channels'
    $idocChannelsSheet = IdocChannelToExcel $idocChannels $idocChannelsSheet
    $idocChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
}

if($restChannels){
    $restChannelsSheet = $workbook.Worksheets.add()
    $restChannelsSheet.name = 'REST Channels'
    $restChannelsSheet = RestChannelToExcel $restChannels $restChannelsSheet
    $restChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
}

if($soapChannels){
    $soapChannelsSheet = $workbook.Worksheets.add()
    $soapChannelsSheet.name = 'SOAP Channels'
    $soapChannelsSheet = SoapChannelToExcel $soapChannels $soapChannelsSheet
    $soapChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
}

if($sftpChannels){
    $sftpChannelSheet = $workbook.Worksheets.add()
    $sftpChannelSheet.name = 'SFTP Channels'
    $sftpChannelSheet = SftpChannelToExcel $sftpChannels $sftpChannelSheet
    $sftpChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
}

if($fileChannels){
    $fileChannelSheet = $workbook.Worksheets.add()
    $fileChannelSheet.Name = 'File Channels'
    $fileChannelSheet = FileChannelToExcel $fileChannels $fileChannelSheet
    $fileChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
}

if($jdbcChannels){
    $jdbcChannelSheet = $workbook.Worksheets.add()
    $jdbcChannelSheet.name = 'JDBC Channels'
    $jdbcChannelSheet = JdbcChannelToExcel $jdbcChannels $jdbcChannelSheet
    $jdbcChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
}

#saving & closing the file
$workbook.SaveAs($outputpath)
$excel.Quit()