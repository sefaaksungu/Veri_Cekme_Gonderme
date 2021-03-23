Attribute VB_Name = "ElementeIslemYaptirma"
' 4.Elementi Aramak ve Varsa Islem Yaptirmak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, element As WebElement
baglan.Start "chrome"
baglan.Get "http://www.milliyet.com.tr/"

Set element = baglan.FindElementByClass("hSoc")

If element Is Nothing Then Exit Sub

MsgBox "elementi buldum"

End Sub
