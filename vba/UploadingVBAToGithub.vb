
Sub SaveVBAScriptTOGit()
    
    'declare variable related to url
    Dim userName As String
    Dim repoName As String
    Dim fileName As String
    Dim accessToken As String
    Dim payload As String

    'declare variable related to HTTP request
    Dim xml_obj As MSXML2.XMLHTTP60
    
    'declare variable related to visual basic editor
    Dim VBAEditor As VBIDE.VBE
    Dim VBProj As VBIDE.vbProject
    Dim vbCodeMod As VBIDE.CodeModule
    Dim vbRawCode As String
    
    'create a reference to VBA editor
    Application.VBE.MainWindow.Visible = True
    Set VBAEditor = Application.VBE
    
    'grab the visual basic project related to my personal macro workbook
    Set VBProj = VBAEditor.VBProjects(4)
    
    'reference a single component in our project and then grab the code module
    Set vbCodeMod = VBProj.VBComponents.Item("ModGit").CodeModule
    
    'grab the raw code in code module
    vbRawCode = vbCodeMod.Lines(startline:=1, Count:=vbCodeMod.CountOfLines)
    
    'base64 encode the string
    Dim RawCodeEncoded As String
    RawCodeEncoded = EncodeBase64(text:=vbRawCode)
    
    Debug.Print RawCodeEncoded
    
    'define our XML HTTP object
    Set xml_obj = New MSXML2.XMLHTTP60
        
        'define URL component
        baseURL = "https://api.github.com/repos/"
        repoName = "VbaProjects/"
        userName = "am235662/"
        fileName = "vba/UploadingVBAToGithub.vb"
        accessToken = "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"
        
        'build the Full URL
        fullurl = baseURL + userName + repoName + "contents/" + fileName + "?ref=master"
        
        'open a new request
        xml_obj.Open bstrmethod:="PUT", bstrurl:=fullurl, varasync:=True
        
        'set the header
        xml_obj.setRequestHeader bstrheader:="Accept", bstrvalue:="application/vnd.github.v3+json"
        xml_obj.setRequestHeader bstrheader:="Authorization", bstrvalue:="token " + accessToken
        
        'define the payload
        payload = "{""message"": ""This is my message"", ""content"":"""
        payload = payload + Application.Clean(RawCodeEncoded)
        payload = payload + """}"
        
        'send the request
        xml_obj.send varbody:=payload
        
        'wait till it is finished
        While xml_obj.readyState <> 4
            DoEvents
        Wend
        
        'print sone info
        Debug.Print "Full URL:" + fullurl
        Debug.Print "Status Text: " + xml_obj.statusText
        Debug.Print "Payload:" + payload
        
End Sub



Function EncodeBase64(text As String) As String
    
    'define our variables
    Dim arrData() As Byte
    Dim objXML As MSXML2.DOMDocument60
    Dim objectNode As MSXML2.IXMLDOMElement
    
    'convert our string to unicode
    arrData = StrConv(text, vbFromUnicode)
    
    'define our dom object
    Set objXML = New MSXML2.DOMDocument60
    Set objectNode = objXML.createElement("b64")
    
    'define data type
    objectNode.DataType = "bin.base64"
    
    'assign the node value
    objectNode.nodeTypedValue = arrData
    
    'return the encoded text
    EncodeBase64 = Replace(objectNode.text, vbLf, "")
    
    'memory clean up
    Set objectNode = Nothing
    Set objXML = Nothing
    
End Function