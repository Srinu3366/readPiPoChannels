function ProcessFileAdapterParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        switch ($AdapterSpecificAttribute.Name) {
            'file.sourceDir' { 
                # Write-Output $AdapterSpecificAttribute.Value
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
            'file.targetFileName' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FileName' -Value $AdapterSpecificAttribute.Value
            }
            Default { }
        }
    }
    return $channelObj
}