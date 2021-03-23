Attribute VB_Name = "FormIslemleri"
' 5.Form Islemleri

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, element As WebElement, coklusecim As SelectElement
baglan.Start "chrome"
baglan.Get "http://compendiumdev.co.uk/selenium/basic_html_form.html"

baglan.FindElementByName("username").SendKeys "mustafa"
baglan.FindElementByXPath("//*[@id='HTMLFormElements']/table/tbody/tr[5]/td/input[2]").Click
baglan.FindElementByXPath("//*[@id='HTMLFormElements']/table/tbody/tr[6]/td/input[3]").Click
baglan.FindElementByName("filename").SendKeys ("C:\Users\mbola\Desktop\deneme.txt")

'For Each element In baglan.FindElementsByXPath("//*[@id='HTMLFormElements']/table/tbody/tr[7]/td/select/option")
'element.Click
'Next element

Set coklusecim = baglan.FindElementByXPath("//*[@id='HTMLFormElements']/table/tbody/tr[7]/td/select").AsSelect
coklusecim.SelectByIndex 0
coklusecim.SelectByIndex 3


End Sub
