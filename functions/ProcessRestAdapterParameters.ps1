function ProcessRestAdapterParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        switch ($AdapterSpecificAttribute.Name) {
            'URLPattern' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'UrlPattern' -Value $AdapterSpecificAttribute.Value
            }
            Default { }
        }
    }    
    return $channelObj
}