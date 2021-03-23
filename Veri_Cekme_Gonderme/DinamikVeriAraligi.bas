Attribute VB_Name = "DinamikVeriAraligi"
' 3.Dinamik Veri Araligindan Veri Çekme(Film Sitesi)

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, resim As Selenium.Image, keys As New keys
baglan.Start "chrome"
baglan.Get "https://www.filmmodu.com/turkce-dublaj"

Dim dd1, dd2 As SelectElement
Set dd1 = baglan.FindElementById("genre").AsSelect
dd1.SelectByIndex 1
Set dd2 = baglan.FindElementById("imdb").AsSelect
dd2.SelectByIndex 3
baglan.FindElementByXPath("/html/body/main/div/div[1]/div/form/div[7]/input").Click

baslamanoktasi:
For i = 1 To 24
resimsayisi = ActiveSheet.Shapes.Count
baglan.FindElementByXPath("/html/body/main/div/div[2]/div[1]/div[" & i & "]/div/img").ScrollIntoView
Set resim = baglan.FindElementByXPath("/html/body/main/div/div[2]/div[1]/div[" & i & "]/div/img").TakeScreenshot
resim.Resize resim.Width * 0.7, resim.Height * 0.7
resim.ToExcel Range("A" & (resimsayisi * 8) - 7)

baglan.FindElementByXPath("/html/body/main/div/div[2]/div[1]/div[" & i & "]/div/a").Click keys.Control
baglan.SwitchToNextWindow

Range("B" & (resimsayisi * 8) - 7) = baglan.FindElementByXPath("/html/body/div[4]/div[1]/div/div[1]/h1").Text
Range("B" & (resimsayisi * 8) - 6) = baglan.FindElementByXPath("/html/body/div[4]/div[2]/div[1]/div[2]/div/div/p[1]").Text

baglan.SwitchToPreviousWindow
baglan.Windows(2).Close
baglan.Windows(1).Activate

Next i

Dim element As WebElement, elementler As WebElements
Set elementler = baglan.FindElementByClass("pagination").FindElementsByTag("a")

For Each element In elementler
If element.Text Like "*Sonraki*" Then
element.Click
GoTo baslamanoktasi
End If
Next element


End Sub
