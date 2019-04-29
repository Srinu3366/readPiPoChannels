function ProcessOtherAdapterTypesParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name $AdapterSpecificAttribute.Name -Value $AdapterSpecificAttribute.Value
    }
    return $channelObj
}