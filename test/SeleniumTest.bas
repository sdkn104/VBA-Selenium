Attribute VB_Name = "SeleniumTest"

Sub test()
    Dim e As SeleniumElement
    
    'Start Selenium
    Dim WebDriver As New SeleniumDriver
    WebDriver.Setup "C:\Users\sdkn1\Desktop\Selenium\chromedriver_win32\chromedriver.exe"
    Application.Wait Now + TimeValue("00:00:02")
    
    'Get HTML Page
    WebDriver.GetUrl "http://htmlpreview.github.io/?https://github.com/sdkn104/VBA-Selenium/blob/master/test/testData.htm"
    'WebDriver.GetUrl "file://C:/Users/sdkn1/Desktop/testData.htm"
    Application.Wait Now + TimeValue("00:00:02")

    'Status
    If Not WebDriver.Status Then Err.Raise 20001

    'Driver: Find Element
    If WebDriver.FindElement("id", "id_text").GetAttribute("id") <> "id_text" Then Err.Raise 20010
    If WebDriver.FindElementById("id_text").GetAttribute("id") <> "id_text" Then Err.Raise 20011
    If WebDriver.FindElementByName("input").GetAttribute("id") <> "id_text" Then Err.Raise 20012
    If WebDriver.FindElementByClassName("c_text").GetAttribute("id") <> "id_text" Then Err.Raise 20013
    If WebDriver.FindElementByTagName("span").GetAttribute("id") <> "id_span" Then Err.Raise 20014
    If WebDriver.FindElementByXpath("//span").GetAttribute("id") <> "id_span" Then Err.Raise 20015
    
    'Element: Find Element, GetAttribute, TagName
    Set e = WebDriver.FindElement("xpath", "/html/body")
    If e.FindElement("id", "id_text").GetAttribute("id") <> "id_text" Then Err.Raise 20020
    If e.FindElementById("id_text").GetAttribute("id") <> "id_text" Then Err.Raise 20021
    If e.FindElementByName("input").GetAttribute("id") <> "id_text" Then Err.Raise 20022
    If e.FindElementByClassName("c_text").GetAttribute("id") <> "id_text" Then Err.Raise 20023
    If e.FindElementByTagName("span").GetAttribute("id") <> "id_span" Then Err.Raise 20024
    If e.FindElementByXpath("//span").GetAttribute("id") <> "id_span" Then Err.Raise 20025
    If Not IsNull(e.FindElementByXpath("//span").GetAttribute("xxxxx")) Then Err.Raise 20026
    If e.FindElementByXpath("//span").TagName <> "span" Then Err.Raise 20027
    
    'Driver: Find Elements
    arr = WebDriver.FindElements("tag name", "input")
    If UBound(arr) - LBound(arr) + 1 <> 3 Then Err.Raise 20031
    For Each v In arr
      If Not v.GetAttribute("class") Like "c_*" Then Err.Raise 20032
    Next
    
    'Element: Find Elements
    arr = WebDriver.FindElementByTagName("body").FindElements("tag name", "input")
    If UBound(arr) - LBound(arr) + 1 <> 3 Then Err.Raise 20035
    For Each v In arr
      If Not v.GetAttribute("class") Like "c_*" Then Err.Raise 20036
    Next
    arr = WebDriver.FindElementByTagName("body").FindElements("tag name", "xxxxx")
    If UBound(arr) >= LBound(arr) Then Err.Raise 20037

    'Send keys, Clear
    WebDriver.FindElement("id", "id_text").SendKeys "abc"
    If WebDriver.FindElement("id", "id_text").GetAttribute("value") <> "abc" Then Err.Raise 20041
    Set e = WebDriver.FindElement("tag name", "form")
    e.FindElement("id", "id_text").SendKeys "def"
    If e.FindElement("id", "id_text").GetAttribute("value") <> "abcdef" Then Err.Raise 20042
    e.FindElement("id", "id_text").Clear
    If e.FindElement("id", "id_text").GetAttribute("value") <> "" Then Err.Raise 20043
    
    'Click, Submit
    WebDriver.FindElement("id", "id_button").Click
    If WebDriver.FindElement("id", "id_text").GetAttribute("value") <> "999" Then Err.Raise 20051
    WebDriver.FindElement("id", "id_submit").Submit
    If WebDriver.FindElement("id", "id_text").GetAttribute("value") <> "111" Then Err.Raise 20052
    
    'ToArray
    tbls = WebDriver.FindElements("xpath", ".//table")
    arr = tbls(1).ToArray
    If arr(1, 1) <> "Firstname" Then Err.Raise 20061
    If Not IsNull(arr(2, 3)) Then Err.Raise 20062
    If arr(3, 4) <> "extra memo" Then Err.Raise 20063
    
    'PageSource
    If Left(WebDriver.PageSource, 15) <> "<!DOCTYPE html>" Then Err.Raise 20071
    
    'Text
    If WebDriver.FindElement("xpath", "/html/body/span").Text <> "button2" Then Err.Raise 20072
    
    'responseText, responseStatus
    s = WebDriver.Status
    If Left(WebDriver.responseText, 1) <> "{" Then Err.Raise 20073
    If WebDriver.responseStatus <> 200 Then Err.Raise 20074
    
    'Error Case
    On Error GoTo OnError
    ExpErrNumber = 10007
    Call WebDriver.FindElement("id", "xxxxxx")
    ExpErrNumber = 10013
    Call WebDriver.FindElement("xxxx", "id_text")
    GoTo EndOnError
OnError:
    ErrNumber = Err.Number
    On Error GoTo 0
    If ErrNumber <> ExpErrNumber Then Err.Raise 20075, "", "ErrNumber:" & ErrNumber & ", expected:" & ExpErrNumber
    On Error GoTo OnError
    Resume Next
EndOnError:
    Err.Description = ""

    'WebDriver.getValueByKey
    If WebDriver.getValueByKey("{""k"":0,{""name"":""value"",""kk"":0}}", "name") <> "value" Then Err.Raise 20081 'String
    If Not IsNull(WebDriver.getValueByKey("{""k"":0,{""name"":null,""kk"":0}}", "name")) Then Err.Raise 20082 'Null
    If WebDriver.getValueByKey("{""k"":0,{""name"":999,""kk"":0}}", "name") <> 999 Then Err.Raise 20083 'number
    If WebDriver.getValueByKey("{""k"":0,{ ""name"" : 999 ,""kk"":0}}", "name") <> 999 Then Err.Raise 20084 'with space
    If WebDriver.getValueByKey("{""k"":0,{ ""name""" & vbTab & vbCrLf & ":" & vbTab & vbCrLf & "999" & vbTab & vbCrLf & ",""kk"":0}}", "name") <> 999 Then Err.Raise 20085 'with tab, cr, if
    'WebDriver.JsonGetValueByKey
    If WebDriver.JsonGetValueByKey("{""k"":0, ""name"" : 999 ,""kk"":0}", "name") <> 999 Then Err.Raise 20085
    If WebDriver.JsonGetValueByKey("{""k"":0, ""name"" : ""abc\t\r\ndef"" ,""kk"":0}", "name") <> "abc" & vbTab & vbCrLf & "def" Then Err.Raise 20086
End Sub
