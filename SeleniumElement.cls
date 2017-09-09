VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeleniumElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Driver As SeleniumDriver
Public elementId As String

' Constructor
Private Sub Class_Initialize()
End Sub


' Destructor
Private Sub Class_Terminate()
   Set Driver = Nothing
End Sub

Public Sub Setup(ByRef WebDriver As SeleniumDriver)
  Set Driver = WebDriver
End Sub


' FIND
Public Function FindElement(ByRef by As String, ByRef value As String) As SeleniumElement
    Dim Elem As SeleniumElement
    Driver.responseStatus = Driver.httpReq("POST", "http://localhost:9515/session/" & Driver.sessionId & "/element/" & elementId & "/element", "{""using"":""" & by & """,""value"":""" & value & """}")
    Set Elem = New SeleniumElement
    Elem.Setup Driver
    Elem.elementId = Driver.getValueByKey(Driver.responseText, "ELEMENT")
    Set FindElement = Elem
    Debug.Print "found elementId: " & Elem.elementId
End Function

' FIND BY XPATH
Public Function FindElementByXpath(ByRef xpath As String) As SeleniumElement
    Set FindElementByXpath = FindElement("xpath", xpath)
End Function

' FIND BY ID
Public Function FindElementById(ByRef id As String) As SeleniumElement
    Set FindElementById = FindElement("id", id)
End Function

' FIND BY NAME
Public Function FindElementByName(ByRef name As String) As SeleniumElement
    Set FindElementByName = FindElement("name", name)
End Function

' FIND BY CLASS NAME
Public Function FindElementByClassName(ByRef className As String) As SeleniumElement
    Set FindElementByClassName = FindElement("class name", className)
End Function


' FIND ELEMENTS
Public Function FindElements(ByRef by As String, ByRef value As String) As Variant
    Dim Elem As SeleniumElement
    Dim arr As Variant, i As Long, elems As Variant
    Driver.responseStatus = Driver.httpReq("POST", "http://localhost:9515/session/" & Driver.sessionId & "/element/" & elementId & "/elements", "{""using"":""" & by & """,""value"":""" & value & """}")
    elems = Driver.getElementIds(Driver.responseText)
    ReDim arr(LBound(elems) To UBound(elems))
    For i = LBound(elems) To UBound(elems)
      Set Elem = New SeleniumElement
      Elem.Setup Driver
      Elem.elementId = elems(i)
      Set arr(i) = Elem
      Set Elem = Nothing
    Next
    FindElements = arr
End Function

' FIND BY TAG NAME
Public Function FindElementByTagName(ByRef tagName As String) As SeleniumElement
    Set FindElementByTagName = FindElement("tag name", tagName)
End Function



' SEND KEYS
Public Function SendKeys(ByRef keys As String) As String
    Driver.responseStatus = Driver.httpReq("POST", "http://localhost:9515/session/" & Driver.sessionId & "/element/" & elementId & "/value", "{""value"":[""" & keys & """]}")
End Function

' SUBMIT
Public Function Submit() As String
    Driver.responseStatus = Driver.httpReq("POST", "http://localhost:9515/session/" & Driver.sessionId & "/element/" & elementId & "/submit")
End Function

' CLICK
Public Function Click() As String
    Driver.responseStatus = Driver.httpReq("POST", "http://localhost:9515/session/" & Driver.sessionId & "/element/" & elementId & "/click")
End Function


' ATTRIBUTE
Public Function GetAttribute(ByRef attributeName As String) As String
    Driver.responseStatus = Driver.httpReq("GET", "http://localhost:9515/session/" & Driver.sessionId & "/element/" & elementId & "/attribute/" & attributeName)
    GetAttribute = getValueByKey(Driver.responseText, "value")
    Debug.Print attributeName & " : " & GetAttribute
End Function

' TEXT
Public Function Text() As String
    Driver.responseStatus = Driver.httpReq("GET", "http://localhost:9515/session/" & Driver.sessionId & "/element/" & elementId & "/text")
    Text = Driver.JsonGetValueByKey(Driver.responseText, "value") '    TODO: unescape JSON String
End Function



'Table to Array   LIMITATION: simple table only. No colspan, rowspan, etc.
Public Function ToArray() As Variant
    Dim rows As Variant, cols As Variant, arr As Variant
    Dim rowCnt As Long, colCnt As Long
    Dim r As Long, c As Long, t As Long
    'Get rows
    rows = FindElements("tag name", "tr")
    If UBound(rows) < 0 Then Exit Function
    rowCnt = UBound(rows) - LBound(rows) + 1
    'Get column count (the max number of columns)
    colCnt = 0
    For r = LBound(rows) To UBound(rows)
      cols = rows(r).FindElements("xpath", ".//td | .//th")
      t = UBound(cols) - LBound(cols) + 1
      If t > colCnt Then colCnt = t
    Next
    If colCnt = 0 Then Exit Function
    'insert into array
    Debug.Print "table: " & rowCnt & " x " & colCnt
    ReDim arr(1 To rowCnt, 1 To colCnt)
    For r = 1 To UBound(arr, 1)
        cols = rows(r - 1 + LBound(rows)).FindElements("xpath", ".//td | .//th")
        For c = 1 To UBound(arr, 2)
            If c <= UBound(cols) - LBound(cols) + 1 Then
              arr(r, c) = cols(c - LBound(cols) - 1).Text
            Else
              arr(r, c) = Null
            End If
        Next
    Next
    ToArray = arr
End Function
