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
            $excelSheet.Cells($i, 2) = 'ComponentID'
            $excelSheet.Cells($i, 3) = 'ChannelID'
            $excelSheet.Cells($i, 4) = 'Direction'
            $excelSheet.Cells($i, 5) = 'JdbcDriver'
            $excelSheet.Cells($i, 6) = 'ConnectionString'
            $excelSheet.Cells($i, 7) = 'UserName'
            $excelSheet.Cells($i, 8) = 'MaximumConcurrency'
            $excelSheet.Cells($i, 9) = 'SqlStatement'
            $i++
        }
        $excelSheet.Cells($i, 1) = $channel.PartyID
        $excelSheet.Cells($i, 2) = $channel.ComponentID
        $excelSheet.Cells($i, 3) = $channel.ChannelID
        $excelSheet.Cells($i, 4) = $channel.Direction
        $excelSheet.Cells($i, 5) = $channel.JdbcDriver
        $excelSheet.Cells($i, 6) = $channel.ConnectionString
        $excelSheet.Cells($i, 7) = $channel.UserName
        $excelSheet.Cells($i, 8) = $channel.MaximumConcurrency
        $excelSheet.Cells($i, 9) = $channel.SqlStatement
        $i++
    }
    return $excelSheet
}