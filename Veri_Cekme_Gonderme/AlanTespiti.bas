Attribute VB_Name = "AlanTespiti"
' 18.Web Sitesinden Veri Alinacak Gonderilecek Alani Tespit Etme Yontemleri

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver
baglan.Start "chrome"
baglan.Get "https://www.facebook.com/"

baglan.FindElementById("u_0_c").SendKeys "merhaba"
baglan.FindElementByName("firstname").SendKeys "selam"
'baglan.FindElementByClass("inputtext _58mg _5dba _2ph-").SendKeys "naber"
baglan.FindElementByXPath("//*[@id='u_0_c']").SendKeys "naber"

End Sub
