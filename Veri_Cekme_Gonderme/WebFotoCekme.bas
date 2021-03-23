Attribute VB_Name = "WebFotoCekme"
' 19.Web Sitesinin Tamamen ya da Kismen Fotografini Cekmek

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, resim As Selenium.Image
baglan.Start "chrome"
baglan.Get "http://uzmanpara.milliyet.com.tr/"

baglan.Window.Maximize
baglan.Wait 500

baglan.FindElementByXPath("/html/body/div[8]/div[6]/div/div[1]/div[3]/div[1]/div").ScrollIntoView
Set resim = baglan.FindElementByXPath("/html/body/div[8]/div[6]/div/div[1]/div[3]/div[1]/div").TakeScreenshot
'resim.ToExcel Range("A1")
resim.SaveAs "C:\Users\mbola\Desktop\New folder\ & resim.png"

End Sub
