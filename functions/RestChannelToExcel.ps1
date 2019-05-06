function RestChannelToExcel {
    param (
        $channels, $excelSheet
    )
    $i = 1
    foreach ($channel in $channels) {
        if ($i -eq 1) {
            $channel | Get-Member -MemberType NoteProperty | ForEach-Object {
                Write-Output $_.Name
            }
            #create the column headers
            $excelSheet.Cells($i, 1) = 'PartyID'
            $excelSheet.Cells($i,1).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 2) = 'ComponentID'
            $excelSheet.Cells($i,2).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 3) = 'ChannelID'
            $excelSheet.Cells($i,3).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 4) = 'Direction'
            $excelSheet.Cells($i,4).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 5) = 'UrlPattern'
            $excelSheet.Cells($i,5).Interior.ColorIndex = 16
            $i++
        }
        $excelSheet.Cells($i, 1) = $channel.PartyID
        $excelSheet.Cells($i, 2) = $channel.ComponentID
        $excelSheet.Cells($i, 3) = $channel.ChannelID
        $excelSheet.Cells($i, 4) = $channel.Direction
        $excelSheet.Cells($i, 5) = $channel.UrlPattern
        $i++
    }
    return $excelSheet
}