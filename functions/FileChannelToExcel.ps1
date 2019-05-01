function FileChannelToExcel {
    param (
        $fileChannels, $fileChannelSheet
    )
    $i = 1
    foreach ($fileChannel in $fileChannels) {
        if ($i -eq 1) {
            $fileChannel | Get-Member -MemberType NoteProperty | ForEach-Object {
                Write-Output $_.Name
            }
            #create the column headers
            $fileChannelSheet.Cells($i, 1) = 'PartyID'
            $fileChannelSheet.Cells($i, 2) = 'ComponentID'
            $fileChannelSheet.Cells($i, 3) = 'ChannelID'
            $fileChannelSheet.Cells($i, 4) = 'Direction'
            $fileChannelSheet.Cells($i, 5) = 'SourceDirectory'
            $fileChannelSheet.Cells($i, 6) = 'TargetDirectory'
            $fileChannelSheet.Cells($i, 7) = 'FileName'
            $i++
        }
        $fileChannelSheet.Cells($i, 1) = $fileChannel.PartyID
        $fileChannelSheet.Cells($i, 2) = $fileChannel.ComponentID
        $fileChannelSheet.Cells($i, 3) = $fileChannel.ChannelID
        $fileChannelSheet.Cells($i, 4) = $fileChannel.Direction
        $fileChannelSheet.Cells($i, 5) = $fileChannel.SourceDirectory
        $fileChannelSheet.Cells($i, 6) = $fileChannel.TargetDirectory
        $fileChannelSheet.Cells($i, 7) = $fileChannel.FileName
        $i++
    }
    return $fileChannelSheet
}