function JdbcChannelToExcel {
    param (
        $channels, $excelSheet
    )
    $i = 1
    foreach ($channel in $channels) {
        if ($i -eq 1) {
            # $fileChannel | Get-Member -MemberType NoteProperty | ForEach-Object{
            #     Write-Output $_.Name
            # }
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
            $excelSheet.Cells($i, 6) = 'JdbcDriver'
            $excelSheet.Cells($i,6).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 7) = 'ConnectionString'
            $excelSheet.Cells($i,7).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 8) = 'UserName'
            $excelSheet.Cells($i,8).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 9) = 'MaximumConcurrency'
            $excelSheet.Cells($i,9).Interior.ColorIndex = 16
            $excelSheet.Cells($i, 10) = 'SqlStatement'
            $excelSheet.Cells($i,10).Interior.ColorIndex = 16
            $i++
        }
        $excelSheet.Cells($i, 1) = $channel.PartyID
        $excelSheet.Cells($i, 2) = $channel.ComponentID
        $excelSheet.Cells($i, 3) = $channel.ChannelID
        $excelSheet.Cells($i, 4) = $channel.Direction
        $excelSheet.Cells($i, 5) = $channel.AdapterEngine
        $excelSheet.Cells($i, 6) = $channel.JdbcDriver
        $excelSheet.Cells($i, 7) = $channel.ConnectionString
        $excelSheet.Cells($i, 8) = $channel.UserName
        $excelSheet.Cells($i, 9) = $channel.MaximumConcurrency
        $excelSheet.Cells($i, 10) = $channel.SqlStatement
        $i++
    }
    return $excelSheet
}