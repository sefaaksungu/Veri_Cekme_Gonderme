Attribute VB_Name = "EtiketBazindaVeriAlma"
' 16.Web Sayfasindaki Verileri Etiket Bazinda Almak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver

baglan.Start "chrome"
baglan.Get "https://www.ntv.com.tr/"

h1etiketi = baglan.FindElementByTag("h1").Text
h2etiketi = baglan.FindElementByTag("h2").Text
petiketi = baglan.FindElementByTag("p").Text

End Sub
