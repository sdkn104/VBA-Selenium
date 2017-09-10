# VBA-Selenium

VBA Class for Selenium WebDriver

now, only a few methods implemented.

### Example

```vb
Dim driver As New SeleniumLib
driver.Setup("C:\......\chromedriver.exe")
driver.GetUrl("http://www.example.com")
```
see [SeleniumTest.bas](/SeleniumTest.bas)


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
