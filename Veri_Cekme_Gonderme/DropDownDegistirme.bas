Attribute VB_Name = "DropDownDegistirme"
' 10.NTV Hava Durumu DropDown Degistirme

Private Sub CommandButton1_Click()
Dim baglan As New Selenium.WebDriver, secim As SelectElement
baglan.Start "chrome"
baglan.Get "https://www.ntv.com.tr/"

baglan.Wait 1500
baglan.FindElementByClass("weather-dropdown").Click

baglan.Wait 1000

Set secim = baglan.FindElementById("weather-cities").AsSelect
secim.SelectByText "Ankara"
baglan.Wait 1000

baglan.FindElementByCss("#homepage > div.weather.boot > div > div > div > div > div.modal-footer > button").Click
                        
End Sub
