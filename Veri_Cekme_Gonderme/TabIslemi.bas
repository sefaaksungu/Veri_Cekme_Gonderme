Attribute VB_Name = "TabIslemi"
' 6.Islemi Geri Almak ve Form Doldururken Tab Islemi Yaptirmak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver
baglan.Start "chrome"
baglan.Get "https://www.kitapyurdu.com/index.php?route=account/register"

baglan.FindElementByName("firstname").SendKeys "Mustafa"

'baglan.FindElementByName("firstname").SendKeys baglan.Keys.Tab

baglan.FindElementByName("firstname").SendKeys baglan.keys.Control, "z"

Stop

End Sub
