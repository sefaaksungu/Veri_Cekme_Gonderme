Attribute VB_Name = "DosyaIndirmeKaydetme"
' 17.Web Sayfasindan Dosya Indirmek ve Kaydetmek

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver

konum = ThisWorkbook.Path
baglan.SetPreference "download.default_directory", konum
baglan.SetPreference "donwload.directory_upgrade", True
baglan.SetPreference "download.prompt_for_download", False

baglan.Start "chrome"
baglan.Get "https://www.bddk.org.tr/BultenGunluk"

baglan.FindElementByXPath("//*[@id='excelTL']").Click

End Sub
