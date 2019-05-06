# readPiPoChannels

This powershell script reads all the channels in PI system and reads the XML of the channel configuration. Then writes the connection parameters of the channel to an excel file.

This will be helpful to generate documentation for all the channels available in the system.

After cloning this repository create a Properties.json file in the root directory of this repository.

Sample properties json looks like 

```javascript
{
    "pi_system_url" : "https://host:port/CommunicationChannelService/HTTPBasicAuth",
    "file_name" : "C:\\PIPOChannels-Test.xlsx"
}
```

Login credentials for `CommunicationChannelService`

A pop-up will be displayed asking for PI system credentials. These credentials will be used to access `CommunicationChannelService`

![Enter PI System credentials](docs\images\Login.png)