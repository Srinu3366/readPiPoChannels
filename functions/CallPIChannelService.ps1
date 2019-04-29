function CallPIChannelService {
    param (
        [Xml] $payload, [string] $authorization, [string] $url
    )
    $soapWebRequest = [System.Net.WebRequest]::Create($url)
    $soapWebRequest.Headers.Add("SOAPAction", "query")
    $soapWebRequest.Headers.Add("Authorization", $authorization)
    $soapWebRequest.ContentType = "text/xml;charset=utf-8"
    $soapWebRequest.Accept = "text/xml"
    $soapWebRequest.Method = "POST"
    $soapWebRequest.UseDefaultCredentials = $true

    #Initiating Send
    $requestStream = $soapWebRequest.GetRequestStream() 
    $payload.Save($requestStream) 
    $requestStream.Close() 

    #Send Complete, Waiting For Response.
    $resp = $soapWebRequest.GetResponse() 
    $responseStream = $resp.GetResponseStream() 
    $soapReader = [System.IO.StreamReader]($responseStream) 
    $ReturnXml = [Xml] $soapReader.ReadToEnd()
    $responseStream.Close() 
    return $ReturnXml
}