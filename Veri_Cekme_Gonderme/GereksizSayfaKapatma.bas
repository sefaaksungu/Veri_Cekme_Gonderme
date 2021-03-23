Attribute VB_Name = "GereksizSayfaKapatma"
' 12.Otomatik Acilan Gereksiz Sayfalari Kapatmak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, anapencere As Selenium.Window
baglan.Start "chrome"
baglan.Get "https://www.ntv.com.tr/"

baglan.Window.Maximize
Set anapencere = baglan.Window
baglan.FindElementByXPath("//*[@id='sticky-wrapper']/nav/div/ul/li[1]/a").Click

For Each pencere In baglan.Windows
If Not pencere.Equals(anapencere) Then
pencere.Close
End If
Next pencere

End Sub
