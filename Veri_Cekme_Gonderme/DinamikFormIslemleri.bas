Attribute VB_Name = "DinamikFormIslemleri"
' 2.Dinamik Form (Hareketli Textbox) Islemleri

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver
baglan.Start "chrome"
baglan.Get "http://demos.codexworld.com/add-remove-input-fields-dynamically-using-jquery/"

For i = 1 To 5
baglan.FindElementByXPath("(//input[@name='field_name[]'])[last()]").SendKeys i
baglan.FindElementByClass("add_button").Click

baglan.Wait 500
Next i

End Sub
