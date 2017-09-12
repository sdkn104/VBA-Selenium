# VBA-Selenium

VBA API for Selenium WebDriver

Now, only limited number of [Selenium commands](https://github.com/SeleniumHQ/selenium/wiki/JsonWireProtocol#command-reference) are implemented.

### Example

```vb
Dim driver As New SeleniumDriver
Dim e As SeleniumElement

driver.Setup("C:\path\to\chromedriver.exe", "chrome")
driver.GetUrl("http://www.example.com")
Set e = driver.FindElementById("id1")
e.Click
Debug.Print e.Text
```
### API List

Class SeleniumDriver

```vb
Public Function Setup(Optional ByRef driverPath As String = "", Optional browser As String = "chrome") As String
Public Function Status() As Boolean
Public Function GetUrl(ByRef url As String) As String
Public Function FindElement(ByRef by As String, ByRef value As String) As SeleniumElement
Public Function FindElementByXpath(ByRef xpath As String) As SeleniumElement
Public Function FindElementByName(ByRef name As String) As SeleniumElement
Public Function FindElementById(ByRef id As String) As SeleniumElement
Public Function FindElementByClassName(ByRef className As String) As SeleniumElement
Public Function FindElementByTagName(ByRef TagName As String) As SeleniumElement
Public Function FindElements(ByRef by As String, ByRef value As String) As Variant
Public Function PageSource() As String
```

Class SeleniumElement

```vb
Public Function FindElement(ByRef by As String, ByRef value As String) As SeleniumElement
Public Function FindElementByXpath(ByRef xpath As String) As SeleniumElement
Public Function FindElementById(ByRef id As String) As SeleniumElement
Public Function FindElementByName(ByRef name As String) As SeleniumElement
Public Function FindElementByClassName(ByRef className As String) As SeleniumElement
Public Function FindElementByTagName(ByRef TagName As String) As SeleniumElement
Public Function FindElements(ByRef by As String, ByRef value As String) As Variant
Public Function SendKeys(ByRef keys As String) As String
Public Function Submit() As String
Public Function Click() As String
Public Function Clear() As String
Public Function GetAttribute(ByRef attributeName As String) As Variant
Public Function Text() As String
Public Function TagName() As String
Public Function ToArray() As Variant
```

### Tested

* Using [SeleniumTest.bas](/SeleniumTest.bas)
* MS Excel 2000, Windows10, [Google Chrome Driver](http://www.seleniumhq.org/download/) 2.32, Google Chrome 61.0.3163.79

### Reference

* Good overview of Selenium
    * (https://app.codegrid.net/entry/selenium-1)
    * (http://codezine.jp/article/detail/10225?p=2)
* Selenium
    * Official (http://www.seleniumhq.org/)
* WebDriver REST API
    * WebDriver API W3C standard (latest, ongoing) (https://www.w3.org/TR/webdriver/)
    * WebDriver Wire Protocol (older but good) (https://github.com/SeleniumHQ/selenium/wiki/JsonWireProtocol)
* Selenium Python
    * not official reference (http://selenium-python.readthedocs.io/index.html)
    * official reference (https://seleniumhq.github.io/selenium/docs/api/py/index.html)

### License

MIT License
