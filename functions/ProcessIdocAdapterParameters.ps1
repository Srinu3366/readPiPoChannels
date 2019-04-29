function ProcessIdocAdapterParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        switch ($AdapterSpecificAttribute.Name) {
            'ResourceAdapterName' { 
                # Write-Output $AdapterSpecificAttribute.Value 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ResourceAdapterName' -Value $AdapterSpecificAttribute.Value
            }
            'MultipleIdocsInIdocXML' { 
                if ($AdapterSpecificAttribute.Value -eq 1) 
                { 
                    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'MultipleIdocsInIdocXML' -Value 'true'
                } 
                else 
                { 
                    Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'MultipleIdocsInIdocXML' -Value 'false'
                } 
            }
            'NumberOfIdocsInIdocXML' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'NumberOfIdocsInXml' -Value $AdapterSpecificAttribute.Value
            }
            Default { }
        }
    }
    return $channelObj
}