function ProcessSoapAdapterParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        switch ($AdapterSpecificAttribute.Name) {
            'XMBWS.TargetURL' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'TargetUrl' -Value $AdapterSpecificAttribute.Value
            }
            'XMBWS.User' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'UserName' -Value $AdapterSpecificAttribute.Value 
            }
            'httpDestination' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'HttpDestination' -Value $AdapterSpecificAttribute.Value
            }
            Default { }
        }
    }
    return $channelObj
}