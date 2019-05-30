function FilechannelToExcel {
    param (
        $channels, $excelSheet
    )
    $i = 1
    foreach ($channel in $channels) {
        if ($i -eq 1) {
            #create the column headers
            $excelSheet.Cells($i, 1) = 'PartyID' 
            $excelSheet.Cells($i,1).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 2) = 'ComponentID'
            $excelSheet.Cells($i,2).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 3) = 'ChannelID'
            $excelSheet.Cells($i,3).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 4) = 'Direction'
            $excelSheet.Cells($i,4).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 5) = 'AdapterEngine'
            $excelSheet.Cells($i,5).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 6) = 'SourceDirectory'
            $excelSheet.Cells($i,6).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 7) = 'TargetDirectory'
            $excelSheet.Cells($i,7).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 8) = 'FileName'
            $excelSheet.Cells($i,8).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 9) = 'FtpHost'
            $excelSheet.Cells($i,9).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 10) = 'FtpPort'
            $excelSheet.Cells($i,10).Interior.ColorIndex = 16
            $i++
        }
        $excelSheet.Cells($i, 1) = $channel.PartyID
        $excelSheet.Cells($i, 2) = $channel.ComponentID
        $excelSheet.Cells($i, 3) = $channel.ChannelID
        $excelSheet.Cells($i, 4) = $channel.Direction
        $excelSheet.Cells($i, 5) = $channel.AdapterEngine
        $excelSheet.Cells($i, 6) = $channel.SourceDirectory
        $excelSheet.Cells($i, 7) = $channel.TargetDirectory
        $excelSheet.Cells($i, 8) = $channel.FileName
        $excelSheet.Cells($i, 9) = $channel.FtpHost
        $excelSheet.Cells($i, 10) = $channel.FtpPort
        $i++
    }
    return $excelSheet
}