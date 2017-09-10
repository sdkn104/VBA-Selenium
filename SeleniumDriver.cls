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
' Selenium API
'
' VBA API for Selenium WebDriver
'
Option Explicit

' Public Member
Public sessionId As String
Public responseText As String
Public responseStatus As String

' Private Member
Private XHR As Object
Private WebDriverPath As String


' Constructor
'  - Setup() should be called after construct
Private Sub Class_Initialize()
    Set XHR = CreateObject("MSXML2.XMLHTTP")
End Sub


' Destructor
Private Sub Class_Terminate()
    'Delete session
    Call httpReq("DELETE", "http://localhost:9515/session/" & sessionId)
    
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


' Setup driver and session
Public Function Setup(Optional ByRef driverPath As String = "") As String
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
    Call httpReq("POST", "http://localhost:9515/session", "{""desiredCapabilities"":{""browser"":""chrome""}}")
    sessionId = getValueByKey(responseText, "sessionId")
    Debug.Print "sessionId: " & sessionId
End Function


' DRIVER STATUS
Public Function Status() As Boolean
    Dim st As String
    Status = False
    Call httpReq("GET", "http://localhost:9515/status")
    If responseStatus < "400" Then Status = True
End Function

' GET
Public Function GetUrl(ByRef url As String) As String
    GetUrl = httpReq("POST", "http://localhost:9515/session/" & sessionId & "/url", "{""url"":""" & url & """}")
End Function


' FIND ELEMENT
Public Function FindElement(ByRef by As String, ByRef value As String) As SeleniumElement
    Call httpReq("POST", "http://localhost:9515/session/" & sessionId & "/element", "{""using"":""" & by & """,""value"":""" & value & """}")
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
    Set FindElementByName = FindElement("name", name)
End Function

' FIND BY ID
Public Function FindElementById(ByRef id As String) As SeleniumElement
    Set FindElementById = FindElement("id", id)
End Function

' FIND BY CLASS NAME
Public Function FindElementByClassName(ByRef className As String) As SeleniumElement
    Set FindElementByClassName = FindElement("class name", className)
End Function


' FIND ELEMENTS
Public Function FindElements(ByRef by As String, ByRef value As String) As Variant
    Dim Elem As SeleniumElement
    Dim arr As Variant, i As Long, elems As Variant
    Call httpReq("POST", "http://localhost:9515/session/" & sessionId & "/elements", "{""using"":""" & by & """,""value"":""" & value & """}")
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
    Call httpReq("GET", "http://localhost:9515/session/" & sessionId & "/source")
    PageSource = JsonGetValueByKey(responseText, "value")
End Function


'------ COMMON MEMBER ----------------------------------------------------------------------------

'  Not real "PUBLIC" members but used also by SeleniumElement class

Public Function httpReq(method As String, url As String, Optional data As String = "") As String
     'XHR may cause fatal error for not-found url etc.
     XHR.Open method, url, False
     XHR.SetRequestHeader "Content-Type", "application/json"
     XHR.Send data
     responseText = XHR.responseText
     responseStatus = XHR.Status
     httpReq = XHR.Status
     
     'Debug.Print "resp " & responseText
     
     If Len(httpReq) <> 3 Or httpReq >= "400" Then
       Err.Raise 10999, "httpReq", "HTTP Error. Response Status = " & httpReq
     End If
End Function


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


' string type("xxx") -> string
' null -> Null
' number type -> numeric (Long, Double, etc)
' true -> Err
' false -> Err
Public Function getValueByKey(json As String, key As String) As Variant
    Dim RE1 As Object, RE2 As Object
    Dim s As String
    Set RE1 = CreateObject("VBScript.RegExp")
    Set RE2 = CreateObject("VBScript.RegExp")
    RE1.Pattern = ".*""" & key & """:([^""},]+)[""},].*"
    RE2.Pattern = ".*""" & key & """:""([^""]*)"".*"
    If RE1.test(json) Then
      s = RE1.Replace(json, "$1")  'TODO: correct handling of JSON syntax
      If s = "null" Then
        getValueByKey = Null
      ElseIf IsNumeric(s) Then
        getValueByKey = Val(s)
      Else
        Err.Raise 10999, "getValueByKey", "illegal response value: " & json
      End If
    ElseIf RE2.test(json) Then
      getValueByKey = RE2.Replace(json, "$1")  'TODO: correct handling of JSON syntax
    Else
      Err.Raise 10999, "getValueByKey", "illegal response value: " & json
    End If
End Function

Public Function getElementIds(json As String) As Variant
    Dim RE As Object
    Dim s As String, v As Variant
    Set RE = CreateObject("VBScript.RegExp")
    RE.Pattern = ".*""value"":\[([^\]]*)\].*"
    s = RE.Replace(json, "$1")  'TODO: correct handling of JSON syntax
    s = Replace(s, """ELEMENT"":", "")
    s = Replace(s, """", "")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")
    v = Split(s, ",")
    getElementIds = v
End Function

