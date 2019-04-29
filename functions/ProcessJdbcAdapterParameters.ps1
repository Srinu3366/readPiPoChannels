function ProcessJdbcAdapterParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        switch ($AdapterSpecificAttribute.Name) {
            'jdbcDriver' { 
                # Write-Output $AdapterSpecificAttribute.Value
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'JdbcDriver' -Value $AdapterSpecificAttribute.Value
            }
            'connectionURL' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'ConnectionString' -Value $AdapterSpecificAttribute.Value
            }
            'dbuser' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'UserName' -Value $AdapterSpecificAttribute.Value
            }
            'maximumConcurrency' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'MaximumConcurrency' -Value $AdapterSpecificAttribute.Value
            }
            'queryStatement' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'SqlStatement' -Value $AdapterSpecificAttribute.Value
            }
            Default { }
        }
    }
    return $channelObj
}