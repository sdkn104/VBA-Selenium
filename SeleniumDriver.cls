VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeleniumDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
' VBA API for Selenium WebDriver
'

Option Explicit

' Public Member
Public sessionId As String
Public responseText As String
Public responseStatus As String
Public driverUrl As String

' Private Member
Private XHR As Object
Private WebDriverPath As String

' Constructor
'  - Setup() should be called after construct
Private Sub Class_Initialize()
    Set XHR = CreateObject("MSXML2.XMLHTTP")
    driverUrl = "http://localhost:9515/"
End Sub


' Destructor
Private Sub Class_Terminate()
    'Delete session
    Call httpReq("DELETE", driverUrl & "session/" & sessionId)
    
    'Kill driver
    Dim WSH As Object, FSO As Object
    Dim nm As String, code As String
    If WebDriverPath <> "" Then
      Set WSH = CreateObject("Wscript.Shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      nm = FSO.GetFile(WebDriverPath).name
      code = WSH.Run("cmd /c Taskkill /F /IM " & nm & "", 0, True)
      Debug.Print "kill driver " & nm & ". status = " & code
    End If

    Set XHR = Nothing
End Sub


' Start driver and create session
Public Function Setup(Optional ByRef driverPath As String = "", Optional browser As String = "chrome") As String
    Dim WSH As Object, FSO As Object
    Dim st As Boolean
    
    'Check WebDriver is running
    On Error Resume Next
    st = Status()
    On Error GoTo 0
    
    'Start driver
    If st = True Then
      Debug.Print "WebDriver is ready. No need to start web driver."
    ElseIf driverPath <> "" Then
      'Start WebDriver
      Set WSH = CreateObject("Wscript.Shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      If Not FSO.FileExists(driverPath) Then Err.Raise 10999, "Setup", "Error: file not found : " & driverPath
      WebDriverPath = driverPath
      Call WSH.Run(driverPath, 0, False) 'non-blocking call
      Debug.Print "starting driver: " & driverPath
      Application.Wait Now + TimeValue("00:00:04") 'wait for driver starting
    Else
      Err.Raise 10999, "Setup", "Error: driver path is not specified while no web driver is running."
    End If
    
    'Get Selenium SessionId
    Call httpReq("POST", driverUrl & "session", "{""desiredCapabilities"":{""browser"":""" & browser & """}}")
    sessionId = getValueByKey(responseText, "sessionId")
    Debug.Print "sessionId: " & sessionId
    Debug.Print "response: " & responseText
End Function


' DRIVER STATUS
Public Function Status() As Boolean
    Status = False
    Call httpReq("GET", driverUrl & "status")
    If responseStatus < "400" Then Status = True
End Function


' GET
Public Function GetUrl(ByRef url As String) As String
    GetUrl = httpReq("POST", driverUrl & "session/" & sessionId & "/url", "{""url"":""" & url & """}")
End Function


' FIND ELEMENT
Public Function FindElement(ByRef by As String, ByRef value As String) As SeleniumElement
    Call httpReq("POST", driverUrl & "session/" & sessionId & "/element", "{""using"":""" & by & """,""value"":""" & value & """}")
    Dim Elem As New SeleniumElement
    Elem.Setup Me
    Elem.elementId = getValueByKey(responseText, "ELEMENT")
    Set FindElement = Elem
    Debug.Print "found elementId: " & Elem.elementId
End Function

' FIND BY XPATH
Public Function FindElementByXpath(ByRef xpath As String) As SeleniumElement
    Set FindElementByXpath = FindElement("xpath", xpath)
End Function

' FIND BY NAME
Public Function FindElementByName(ByRef name As String) As SeleniumElement
    Set FindElementByName = FindElement("xpath", "//*[@name='" & name & "']")
End Function

' FIND BY ID
Public Function FindElementById(ByRef id As String) As SeleniumElement
    Set FindElementById = FindElement("xpath", "//*[@id='" & id & "']")
End Function

' FIND BY CLASS NAME
Public Function FindElementByClassName(ByRef className As String) As SeleniumElement
    Set FindElementByClassName = FindElement("xpath", "//*[@class='" & className & "']")
End Function

' FIND BY TAG NAME
Public Function FindElementByTagName(ByRef TagName As String) As SeleniumElement
    Set FindElementByTagName = FindElement("tag name", TagName)
End Function


' FIND ELEMENTS
Public Function FindElements(ByRef by As String, ByRef value As String) As Variant
    Dim Elem As SeleniumElement
    Dim arr As Variant, i As Long, elems As Variant
    Call httpReq("POST", driverUrl & "session/" & sessionId & "/elements", "{""using"":""" & by & """,""value"":""" & value & """}")
    elems = getElementIds(responseText)
    If UBound(elems) >= LBound(elems) Then
      ReDim arr(LBound(elems) To UBound(elems))
    Else
      arr = Split("", ",") 'empty String array
    End If
    For i = LBound(elems) To UBound(elems)
      Set Elem = New SeleniumElement
      Elem.Setup Me
      Elem.elementId = elems(i)
      Set arr(i) = Elem
      Set Elem = Nothing
    Next
    FindElements = arr
End Function


' PAGE SOURCE
Public Function PageSource() As String
    Call httpReq("GET", driverUrl & "session/" & sessionId & "/source")
    PageSource = JsonGetValueByKey(responseText, "value")
End Function


'------ COMMON MEMBER (used also by SeleniumElement class) -------------------------------------------------

' Set Function/Sub as Public, since they are used also by SeleniumElement class

' HTTP request to WebDriver server. Checking command status.
Public Function httpReq(method As String, url As String, Optional data As String = "") As String
     Dim st As Variant
     'XHR may cause fatal error for not-found url etc.
     XHR.Open method, url, False
     XHR.SetRequestHeader "Content-Type", "application/json"
     XHR.Send data
     responseText = XHR.responseText
     responseStatus = XHR.Status
     httpReq = XHR.Status
     
     Debug.Print "resp " & responseText
     
     'Invalid request. response body shall be text/plain.
     If XHR.Status Like "4??" Then
       Err.Raise 10000 + XHR.Status, "httpReq", "HTTP Respond Error (Status = " & XHR.Status & ") " & XHR.responseText & vbCrLf & url & vbCrLf & data
     End If
    
     'get command status
     st = getValueByKey(responseText, "status")

     'Command fail
     If IsNumeric(st) And st <> 0 Then
       Err.Raise 10000 + st, "httpReq", "WebDriver command failed (Status = " & st & ") " & XHR.responseText & vbCrLf & url & vbCrLf & data
     End If
     
     'Usually not reached here.  5xx status should be command fail
     If XHR.Status Like "5??" Then
       Err.Raise 10000 + XHR.Status, "httpReq", "HTTP Respond Error (Status = " & XHR.Status & ") " & XHR.responseText & vbCrLf & url & vbCrLf & data
     End If
End Function


' Json parser.
'   - get value of name(key) at top level
' have to add a reference to "Microsoft Script Control 1.0".
Public Function JsonGetValueByKey(jsonString As String, key As String) As String
    ' from https://stackoverflow.com/questions/6627652/parsing-json-in-excel-vba/7300926
    'TODO: this code is robust??? has limitation??
    Dim ScriptEngine As ScriptControl
    Dim JsonObj As Object
    Set ScriptEngine = New ScriptControl
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "Object.prototype.myitem=function( i ) { return this[i] } ; "
    Set JsonObj = ScriptEngine.Eval("(" + jsonString + ")")
    JsonGetValueByKey = JsonObj.myitem(key)
End Function


' Poor Json parser
'   - get value of name(key) at any level. if there are multiple entry of the name, it pick up the first one.
'   - the found value must be null, number, simple string (not containing escaped double-quote)
Public Function getValueByKey(json As String, key As String) As Variant 'TODO: correct handling of JSON syntax
    Dim RE1 As Object, RE2 As Object, RE3 As Object
    Dim s As String
    Set RE1 = CreateObject("VBScript.RegExp")
    Set RE2 = CreateObject("VBScript.RegExp")
    Set RE3 = CreateObject("VBScript.RegExp")
    RE1.Pattern = ".*""" & key & """[ \t\r\n]*:[ \t\r\n]*([^""{\[, \t\r\n]+)[ \t\r\n]*[},].*"
    RE2.Pattern = ".*""" & key & """[ \t\r\n]*:[ \t\r\n]*""([^""]*)"".*"
    RE3.Pattern = """" & key & """[ \t\r\n]*:"
    If Not RE3.test(json) Then '--- key not found
      getValueByKey = "the key not found"
    ElseIf RE1.test(json) Then '--- value is other than string("...), array([...), nor object({...)
      s = RE1.Replace(json, "$1")
      If s = "null" Then
        getValueByKey = Null
      ElseIf IsNumeric(s) Then
        getValueByKey = Val(s)
      Else
        Err.Raise 10999, "getValueByKey", "illegal response value: " & json
      End If
    ElseIf RE2.test(json) Then 'value is string
      getValueByKey = RE2.Replace(json, "$1")  'TODO: correct handling of JSON syntax
    Else 'value is object or array
      getValueByKey = json
    End If
End Function


' get element ids from JSON string like "value":["ELEMENT":"idxxxx","ELEMENT":"idxxxx","ELEMENT":"idxxxx"]
Public Function getElementIds(json As String) As Variant
    Dim RE As Object
    Dim s As String, v As Variant
    Set RE = CreateObject("VBScript.RegExp")
    RE.Pattern = ".*""value""[ \t\r\n]*:[ \t\r\n]*\[[ \t\r\n]*([^\[ \t\r\n]*)[ \t\r\n]*\].*"
    s = RE.Replace(json, "$1")  'TODO: correct handling of JSON syntax
    s = Replace(s, """ELEMENT"":", "")
    s = Replace(s, """", "")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")
    v = Split(s, ",")
    getElementIds = v
End Function


