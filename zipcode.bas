Attribute VB_Name = "CountyLookup"

Function getcounty(zip As String) As String
    Dim strUrl As String    ' Our URL which will include the authentication info
    Dim strReq As String    ' The body of the POST request
    
    ' Make sure to reference the "Microsoft XML, v6.0" library (Tools -> References).
    ' You can use an older version, too, just be sure to change the number in this next line.
    Dim xmlHttp As New MSXML2.XMLHTTP60
       
    Dim key As String
    key = "GOOGLE_API_KEY"
    
    strUrl = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & zip & "&key=" & key
    
    ' Perform the request
    With xmlHttp
        .Open "GET", strUrl, False                      ' Prepare GET request
        .setRequestHeader "Content-Type", "text/xml"    ' Sending XML ...
        .setRequestHeader "Accept", "text/xml"          ' ... expect XML in return.
        .Send ""                                        ' Send request body
    End With
    
    
    ' The request has been saved into xmlHttp.responseText and is
    ' now ready to be parsed. Remember that fields in our XML response may
    ' change or be added to later, so make sure your method of parsing accepts that.
    ' Google and Stack Overflow are replete with helpful examples.
    
    Dim xmlDoc As MSXML2.DOMDocument            ' In Office 2010, use: MSXML2.DOMDocument60
    Set xmlDoc = New MSXML2.DOMDocument         ' on both these lines

    If Not xmlDoc.LoadXML(xmlHttp.responseText) Then
        Err.Raise xmlDoc.parseError.ErrorCode, , xmlDoc.parseError.reason
    End If
    
    Set county = xmlDoc.selectNodes("//long_name[text()[contains(.,'County')]]")
    Set city = xmlDoc.selectNodes("//type[text()[contains(.,'locality')]]/../long_name")
    Set state = xmlDoc.selectNodes("//type[text()[contains(.,'administrative_area_level_1')]]/../long_name")
    
    If county.length > 0 Then
        getcounty = city.nextnode.text & ", " & county.nextnode.text & ", " & state.nextnode.text
    Else
        getcounty = city.nextnode.text & ", " & state.nextnode.text
    End If

End Function
