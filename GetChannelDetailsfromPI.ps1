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

foreach ($property in Get-Content .\Properties.json | ConvertFrom-Json) {
    $propertyDetail = $property | Get-Member -MemberType NoteProperty
    $name = $propertyDetail.Name
    switch ($name) {
        'pi_system_url' { $url = $property."$name" }
        Default { }
    }
}

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

Write-Host 'Reading Channel XML configuration.....'
foreach ( $channel in $channelObjList ) {
    $request = [Xml] $SOAPReadChannelRequest.OuterXml.Replace("{{PartyID}}", $channel.PartyID).Replace("{{ComponentID}}", $channel.ComponentID).Replace("{{ChannelID}}", $channel.ChannelID)

    $ReturnXml = CallPIChannelService $request $authorization $url
    $direction = $ReturnXml.GetElementsByTagName('rn3:CommunicationChannel').Direction

    $adapterType = $ReturnXml.GetElementsByTagName('rn3:AdapterMetadata').name
    $channelObj = New-Object -TypeName PSObject
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'PartyID' -Value $channel.PartyID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ComponentID' -Value $channel.ComponentID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ChannelID' -Value $channel.ChannelID
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'Direction' -Value $direction
    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'AdapterType' -Value $adapterType

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
$excel.visible = $True
$workbook = $excel.Workbooks.Add()

#give the remaining worksheet a name
$otherChannelsSheet = $workbook.Worksheets.Item(1)
$otherChannelsSheet.name = 'Other Channels'
$otherChannelsSheet = OtherAdapterChannelsToExcel $otherChannels $otherChannelsSheet
#adjusting the column width so all data's properly visible
$otherChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

$idocChannelsSheet = $workbook.Worksheets.add()
$idocChannelsSheet.name = 'Idoc Channels'
$idocChannelsSheet = IdocChannelToExcel $idocChannels $idocChannelsSheet
$idocChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

$restChannelsSheet = $workbook.Worksheets.add()
$restChannelsSheet.name = 'REST Channels'
$restChannelsSheet = RestChannelToExcel $restChannels $restChannelsSheet
$restChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

$soapChannelsSheet = $workbook.Worksheets.add()
$soapChannelsSheet.name = 'SOAP Channels'
$soapChannelsSheet = SoapChannelToExcel $soapChannels $soapChannelsSheet
$soapChannelsSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

$sftpChannelSheet = $workbook.Worksheets.add()
$sftpChannelSheet.name = 'SFTP Channels'
$sftpChannelSheet = SftpChannelToExcel $sftpChannels $sftpChannelSheet
$sftpChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

$fileChannelSheet = $workbook.Worksheets.add()
$fileChannelSheet.Name = 'File Channels'
$fileChannelSheet = FileChannelToExcel $fileChannels $fileChannelSheet
$fileChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null


$jdbcChannelSheet = $workbook.Worksheets.add()
$jdbcChannelSheet.name = 'JDBC Channels'
$jdbcChannelSheet = JdbcChannelToExcel $jdbcChannels $jdbcChannelSheet
$jdbcChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null


#saving & closing the file
$outputpath = join-path -Path $env:USERPROFILE -ChildPath "desktop\PIPOChannelDetails-XP4.xlsx"
$workbook.SaveAs($outputpath)
$excel.Quit()