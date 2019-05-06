function ProcessFileAdapterParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        switch ($AdapterSpecificAttribute.Name) {
            'file.sourceDir' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'SourceDirectory' -Value $AdapterSpecificAttribute.Value
            }
            'ftp.sourceDir' {
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'SourceDirectory' -Value $AdapterSpecificAttribute.Value
            }
            'file.sourceFileName' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FileName' -Value $AdapterSpecificAttribute.Value
            }
            'file.pollInterval' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'PollingInterval' -Value $AdapterSpecificAttribute.Value
            }
            'file.targetDir' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'TargetDirectory' -Value $AdapterSpecificAttribute.Value
            }
            'ftp.targetDir' {
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'TargetDirectory' -Value $AdapterSpecificAttribute.Value
            }
            'file.targetFileName' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FileName' -Value $AdapterSpecificAttribute.Value
            }
            'ftp.host' {
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FtpHost' -Value $AdapterSpecificAttribute.Value
            }
            'ftp.port' {
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FtpPort' -Value $AdapterSpecificAttribute.Value
            }
            Default { }
        }
    }
    return $channelObj
}