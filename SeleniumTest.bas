Attribute VB_Name = "SeleniumTest"

Sub test()
    Dim e As SeleniumElement
    
    'Start Selenium
    Dim WebDriver As New SeleniumDriver
    WebDriver.Setup "C:\Users\sdkn1\Desktop\Selenium\chromedriver_win32\chromedriver.exe"
    Application.Wait Now + TimeValue("00:00:02")

    WebDriver.GetUrl "http://htmlpreview.github.io/?https://github.com/sdkn104/VBA-Selenium/blob/master/test/testData.htm"
    'WebDriver.GetUrl "file://C:/Users/sdkn1/Desktop/testData.htm"
    Application.Wait Now + TimeValue("00:00:02")

    'Find Element
    WebDriver.FindElement("id", "id_text").SendKeys "abcde"
    WebDriver.FindElement("id", "id_button").Click
    'WebDriver.FindElement("id", "id_submit").Submit
    Set e = WebDriver.FindElement("tag name", "form")
    e.FindElement("id", "id_text").SendKeys "abcde"
    e.FindElement("id", "id_button").Click
    'e.FindElement("id", "id_submit").Submit
    
    tbls = WebDriver.FindElements("xpath", "//table")
    tbls = WebDriver.FindElement("xpath", "/html/body").FindElements("xpath", "//table")
    
    tbl1arr = tbls(1).ToArray
    'ActiveSheet.Range("A1").Resize(UBound(tbl1arr, 1), UBound(tbl1arr, 2)) = tbl1arr
    
    Debug.Print WebDriver.FindElement("id", "id_text").GetAttribute("name")
    Debug.Print WebDriver.Status
    Debug.Print Left(WebDriver.PageSource, 100)
    Debug.Print WebDriver.FindElement("xpath", "/html/body/span").Text
End Sub
