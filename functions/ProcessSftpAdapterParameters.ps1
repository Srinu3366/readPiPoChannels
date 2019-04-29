function ProcessSftpAdapterParameters {
    param (
        $adapterAttributes, $channelObj
    )
    foreach ($AdapterSpecificAttribute in $adapterAttributes) {
        switch ($AdapterSpecificAttribute.Name) {
            'serverhost' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'HostName' -Value $AdapterSpecificAttribute.Value
            }
            'serverport' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'Port' -Value $AdapterSpecificAttribute.Value
            }
            'authMethod' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'AuthenticationMethod' -Value $AdapterSpecificAttribute.Value
            }
            'userName' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'UserName' -Value $AdapterSpecificAttribute.Value
            }
            'filePath' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FilePath' -Value $AdapterSpecificAttribute.Value
            }
            'fileName' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FileName' -Value $AdapterSpecificAttribute.Value
            }
            'fileDirectory' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FileDirectory' -Value $AdapterSpecificAttribute.Value
            }
            'regFileName' { 
                Add-Member -InputObject $channelObj -MemberType 'NoteProperty' -Name 'FileName' -Value $AdapterSpecificAttribute.Value
            }
            Default { }
        }
    }
    return $channelObj  
}