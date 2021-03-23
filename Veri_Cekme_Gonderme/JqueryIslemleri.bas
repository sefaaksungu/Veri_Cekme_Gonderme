Attribute VB_Name = "JqueryIslemleri"
' 7.Jquery(Uyari) Mesajlarini Atlatmak ve Doldurmak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver
baglan.Start "chrome"
baglan.Get "http://bootboxjs.com/examples.html"

'baglan.FindElementByXPath("//*[@id='bb-alert-examples']/div/ul/li[1]/p[2]/button").Click
'baglan.SendKeys baglan.Keys.Enter

'baglan.FindElementByXPath("//*[@id='bb-confirm-examples']/div/ul/li[1]/p[2]/button").Click
'baglan.SendKeys baglan.Keys.Escape

baglan.FindElementByXPath("//*[@id='bb-prompt-examples']/div/ul/li[1]/p[2]/button").Click
baglan.FindElementByXPath("/html/body/div[4]/div/div/div[2]/div/form/input").SendKeys "Merhaba!"
baglan.SendKeys baglan.keys.Enter
End Sub
