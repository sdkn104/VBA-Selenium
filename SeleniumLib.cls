VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeleniumLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Selenium API
'
' VBA API for Selenium WebDriver
'
'

Public sessionId As String
Public responseText As String
Public responseStatus As String

Private XHR As Object
Private WebDriverPath As String


' Constructor
'  - Setup() should be called after construct
Private Sub Class_Initialize()
    Set XHR = CreateObject("MSXML2.XMLHTTP")
End Sub


' Destructor
Private Sub Class_Terminate()
    'Close session
    responseStatus = httpReq("DELETE", "http://localhost:9515/session/" & sessionId) 'delete session
    
    'Kill driver
    Dim WSH As Object, FSO As Object
    Dim nm As String, code As String
    If WebDriverPath <> "" Then
      Set WSH = CreateObject("Wscript.Shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      nm = FSO.GetFile(WebDriverPath).Name
      code = WSH.Run("cmd /c Taskkill /F /IM " & nm & "", 0, True)
      Debug.Print "kill driver " & nm & ". status = " & code
    End If
End Sub


' Setup driver and session
Public Function Setup(Optional ByRef driverPath As String = "") As String
    Dim WSH As Object, FSO As Object
    Dim st As Boolean
    
    'Check WebDriver running
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
    responseStatus = httpReq("POST", "http://localhost:9515/session", "{""desiredCapabilities"":{""browser"":""chrome""}}")
    sessionId = getValueByKey(responseText, "sessionId")
    Debug.Print "sessionId: " & sessionId
End Function


' DRIVER STATUS
Public Function Status() As Boolean
    Dim st As String
    Status = False
    responseStatus = httpReq("GET", "http://localhost:9515/status")
    If responseStatus < "400" Then Status = True
End Function

' GET
Public Function GetReq(ByRef url As String) As String
    GetReq = httpReq("POST", "http://localhost:9515/session/" & sessionId & "/url", "{""url"":""" & url & """}")
End Function


' FIND BY XPATH
Public Function FindElementByXpath(ByRef xpath As String) As String
    responseStatus = httpReq("POST", "http://localhost:9515/session/" & sessionId & "/element", "{""using"":""xpath"",""value"":""//span[contains(text(), 'Download Data')]/parent::a""}")
    FindElementByXpath = getValueByKey(responseText, "ELEMENT")
    Debug.Print "elementId: " & FindElementByXpath
End Function

' ATTRIBUTE
Public Function GetAttribute(ByRef elementId As String, ByRef attributeName As String) As String
    responseStatus = httpReq("GET", "http://localhost:9515/session/" & sessionId & "/element/" & elementId & "/attribute/" & attributeName)
    GetAttribute = getValueByKey(responseText, "value")
    Debug.Print attributeName & " : " & GetAttribute
End Function


'------ PRIVATE MEMBER ------
Private Function getValueByKey(json As String, key As String) As String
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    RE.Pattern = ".*""" & key & """:""([^""]*)"".*"
    getValueByKey = RE.Replace(json, "$1")  'TODO: correct handling of JSON syntax
End Function


Private Function httpReq(method As String, url As String, Optional data As String = "") As String
     'On Error GoTo AFTER
     XHR.Open method, url, False
     XHR.SetRequestHeader "Content-Type", "application/json"
     XHR.Send data
     responseText = XHR.responseText
     responseStatus = XHR.Status
     httpReq = XHR.Status
'AFTER:
'     If Len(httpReq) <> 3 Then httpReq = 999 ' not foundなどでもエラーがでるので、エラーでもstatusを設定する。
'     If Err.Number <> 0 Then
'       Err.Raise Err.Number, Err.Source, Err.Description
'     End If
'     On Error GoTo 0
'
'     If Err.Number = -2146697211 Then '指定されたリソースが見つかりません
'     End If
     
     If httpReq < "400" Then
     Else
       Err.Raise 10999, "httpReq", "HTTP Error. Response Status = " & httpReq
     End If
     Debug.Print responseText
End Function
