Attribute VB_Name = "KosulaGoreTiklama"
' 8.Kosula Göre a Etiketine Tiklamak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, element As WebElement, elementler As WebElements, linkler As List

baglan.Start "chrome"
baglan.Get "https://www.ntv.com.tr/"

Set elementler = baglan.FindElementsByTag("a")
Set linkler = elementler.Attribute("href")
linkler.Distinct
linkler.Sort
linkler.ToExcel Range("A1")

linksayisi = linkler.Count

For Each element In elementler
metin = element.Text
If metin Like "*SON*" Then
element.Click
Exit For
End If
Next element

End Sub
