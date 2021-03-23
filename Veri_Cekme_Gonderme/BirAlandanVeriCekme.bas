Attribute VB_Name = "BirAlandanVeriCekme"
'1. Belirli Bir Alandan Veri Çekmek

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, linkler As List
baglan.Start "chrome"
baglan.Get "http://www.milliyet.com.tr/"

Set linkler = baglan.FindElementByClass("oneCikanlar").FindElementsByTag("a").Attribute("href")

End Sub
