Attribute VB_Name = "KelimeCeviriProgrami"
' 11.Online Kelime Ceviri Programi Yazma

Private Sub CommandButton1_Click()
Me.Hide
Dim baglan As New Selenium.WebDriver, listeler As List
baglan.Start "chrome"

If Me.OptionButton1 = True Then
baglan.Get "https://tr.bab.la/sozluk/ingilizce-turkce/"
Else
baglan.Get "https://tr.bab.la/sozluk/turkce-ingilizce/"
End If

For X = 1 To Cells(Rows.Count, "a").End(xlUp).Row
metin = Cells(X, 1)

baglan.FindElementById("bablasearch").SendKeys Cells(X, 1)
baglan.SendKeys baglan.keys.Enter

Set listeler = baglan.FindElementByXPath("/html/body/main/div/div/div/div[1]/div/div[2]/div[2]/ul").FindElementsByTag("li").Text
saydir = listeler.Count

For i = 1 To saydir
If Cells(X, 2) <> "" Then
Cells(X, 2) = Cells(X, 2) & "," & listeler.Item(i)
Else
Cells(X, 2) = listeler.Item(i)
End If
Next i
Next X
End Sub
