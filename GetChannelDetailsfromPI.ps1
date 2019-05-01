$xlAutomatic = -4105
$xlBottom = -4107
$xlCenter = -4108
$xlContext = -5002
$xlContinuous = 1
$xlDiagonalDown = 5
$xlDiagonalUp = 6
$xlEdgeBottom = 9
$xlEdgeLeft = 7
$xlEdgeRight = 10
$xlEdgeTop = 8
$xlInsideHorizontal = 12
$xlInsideVertical = 11
$xlNone = -4142
$xlThin = 2
$xlMedium = -4138
$xlThick = 4 
function FormatExcel {
    param (
        $excelSheet
    )   
    #adjusting the column width so all data's properly visible
    # $usedRange = $fileChannelSheet.UsedRange	
    $excelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
    $selection = $excelSheet.UsedRange
    [void]$selection.select()
    $selection.Borders.Item($xlEdgeLeft).LineStyle = $xlContinuous
    $selection.Borders.Item($xlEdgeLeft).ColorIndex = $xlAutomatic
    $selection.Borders.Item($xlEdgeLeft).Color = 1
    $selection.Borders.Item($xlEdgeLeft).Weight = $xlMedium
    $selection.Borders.Item($xlEdgeTop).LineStyle = $xlContinuous
    $selection.Borders.Item($xlEdgeBottom).LineStyle = $xlContinuous
    $selection.Borders.Item($xlEdgeRight).LineStyle = $xlContinuous
    $selection.Borders.Item($xlInsideVertical).LineStyle = $xlContinuous
    $selection.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous

    return $excelSheet
}

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

Write-Host 'Read Channel Names.....'
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

Write-Host 'Read Channel XML configuration.....'
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

Write-Host 'Prepare Excel Output...'

#open excell
$excel = New-Object -ComObject excel.application
$excel.visible = $True

#add a default workbook
$workbook = $excel.Workbooks.Add()

#give the remaining worksheet a name
$fileChannelSheet = $workbook.Worksheets.Item(1)
$fileChannelSheet.Name = 'File Channels'

$fileChannelSheet = FileChannelToExcel $fileChannels $fileChannelSheet
#adjusting the column width so all data's properly visible
$fileChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null


$jdbcChannelSheet = $workbook.Worksheets.add()
$jdbcChannelSheet.name = 'JDBC Channels'

$jdbcChannelSheet = JdbcChannelToExcel $jdbcChannels $jdbcChannelSheet

$jdbcChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null


$sftpChannelSheet = $workbook.Worksheets.add()
$sftpChannelSheet.name = 'SFTP Channels'

$sftpChannelSheet = SftpChannelToExcel $sftpChannels $sftpChannelSheet

$sftpChannelSheet.UsedRange.EntireColumn.AutoFit() | Out-Null


$soapChannelsSheet = $workbook.Worksheets.add()
$soapChannelsSheet.name = 'SOAP Channels'

$restChannelsSheet = $workbook.Worksheets.add()
$restChannelsSheet.name = 'REST Channels'

$idocChannelsSheet = $workbook.Worksheets.add()
$idocChannelsSheet.name = 'Idoc Channels'

$otherChannelsSheet = $workbook.Worksheets.add()
$otherChannelsSheet.name = 'Other Channels'

#saving & closing the file
$outputpath = join-path -Path $env:USERPROFILE -ChildPath "desktop\PIPOChannelDetails.xlsx"
$workbook.SaveAs($outputpath)
$excel.Quit()