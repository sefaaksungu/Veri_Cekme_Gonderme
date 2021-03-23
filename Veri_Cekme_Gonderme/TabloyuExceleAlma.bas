Attribute VB_Name = "TabloyuExceleAlma"
' 13.Sayfada Yer Alan Tabloyu Excele Aktarmak

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, tablom As TableElement
baglan.Start "chrome"
baglan.Get "http://bymmb.com/ms-excel-fonksiyonlari-turkce-karsiliklari/"

Set tablom = baglan.FindElementByXPath("//*[@id='page']/div/div/div/section/div[2]/article/div/div[2]/div[1]/table").AsTable
tablom.ToExcel Range("A1")

End Sub
