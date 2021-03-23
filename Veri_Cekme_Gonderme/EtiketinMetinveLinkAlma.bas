Attribute VB_Name = "EtiketinMetinveLinkAlma"
' 15.Sayfadaki a Etiketinin Metinlerini ve Linklerini Almak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, elementler As WebElements, metinler As List, uzantilar As List
baglan.Start "chrome"
baglan.Get "http://www.milliyet.com.tr/"

Set elementler = baglan.FindElementsByTag("a")
Set metinler = elementler.Text
Set uzantilar = elementler.Attribute("href")

metinler.Distinct
metinler.Sort
metinler.ToExcel Range("A2")

uzantilar.Distinct
uzantilar.Sort
uzantilar.ToExcel Range("B2")

baglan.Get uzantilar(50)
End Sub
