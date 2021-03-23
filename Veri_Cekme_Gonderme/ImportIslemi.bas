Attribute VB_Name = "ImportIslemi"
' 9.Liste Elementlerinin Tamamini Import Etmek ve Kosula Göre Tiklamak

Sub lietiketleri()
Dim baglan As New Selenium.WebDriver, liler As WebElements, li As WebElement

baglan.Start "chrome"
baglan.Get "https://www.ntv.com.tr/"

Set liler = baglan.FindElementsByTag("li")

For Each li In liler
metin = li.Text
If metin Like "*BIST*" Then
li.Click
Exit For
End If
Next li

End Sub
