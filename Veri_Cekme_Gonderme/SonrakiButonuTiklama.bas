Attribute VB_Name = "SonrakiButonuTiklama"
' 14.Sayfadaki Sonraki Butonuna Tiklama

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, element As WebElement, elementler As WebElements
baglan.Start "chrome"
baglan.Get "https://www.filmmodu.com/en-cok-izlenen-filmler"

baslamanoktasi:
Set elementler = baglan.FindElementsByTag("a")
For Each element In elementler
If element.Text Like "*Sonraki*" Then
element.Click
baglan.FindElementByXPath("/html/body/main/div[3]").ScrollIntoView
baglan.Wait 500
GoTo baslamanoktasi
End If
Next element

End Sub
