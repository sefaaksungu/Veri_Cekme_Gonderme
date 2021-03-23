Attribute VB_Name = "SiparisOlusturma"
Sub satis_siparis()
Dim baglan As New Selenium.WebDriver

baglan.Start "chrome"
baglan.Get "http://**********/login"
baglan.Window.Maximize

baglan.FindElementById("*****").SendKeys "*****"
baglan.FindElementById("*****").SendKeys "*****"
baglan.FindElementById("*****").Click

baglan.Wait 4000
n = 2
k = 2
h = 2
m = 30
For r = 1 To 1000
    
    For j = h To 1000
        baglan.Wait 4000
        baglan.Get "http://*****/add"
        baglan.Wait 4000

        baglan.FindElementByXPath("//*[@id='*****']/div/*****/button/span[2]/span").Click 'müsteri seçimi
        baglan.Wait 4000
        baglan.FindElementByXPath("//*[@id='*****']/div/*****/div[2]/input").SendKeys Cells(j, 1).Value 'müsteri seçimi
        baglan.Wait 4000
        baglan.FindElementByXPath("//*[@id='*****']/div/*****/ul/li[1]/a").Click 'müsteri seçimi
        baglan.Wait 4000
        baglan.WaitForScript "!jQuery.active"
        baglan.ExecuteScript ("$('#*****').select2('val',*****).change()") 'vade seçimi
        'baglan.FindElementByXPath("//*[@id='*****']/a/span[2]").Click 'vade seçimi
        'baglan.Wait 4000
        'baglan.FindElementByXPath("//*[@id='*****']").SendKeys Cells(j, 5).Value 'vade seçimi
        'baglan.Wait 4000
        'baglan.FindElementByXPath("//*[@id='*****']/li").Click 'vade seçimi
        'baglan.Wait 4000
        baglan.ExecuteScript ("$('button[data-id=\'*****']').click()")
        baglan.ExecuteScript ("$('button[data-id=\'*****']').next('div.dropdown-menu.open').find('*****').val('" & Cells(j, 6).Value & "').keyup()")
        baglan.WaitForScript "!jQuery.active"
        baglan.Wait 4000
        baglan.FindElementByXPath("//*[@id='*****']/div/*****/ul/li[1]/a").Click 'satis personeli seçimi
        'baglan.FindElementByXPath("//*[@id='*****']/div/*****/button/span[2]/span").Click 'satis personeli seçimi
        baglan.Wait 4000
        'baglan.FindElementByXPath("//*[@id='*****']/div/*****/div[2]/input").SendKeys Cells(j, 6).Value 'satis personeli seçimi
        'baglan.Wait 4000
        'baglan.FindElementByXPath("//*[@id='*****']/div/*****/ul/li[1]/a").Click 'satis personeli seçimi
        'baglan.Wait 4000
        baglan.ExecuteScript ("$('#*****').val(" & Cells(j, 7) & ").change()") 'istenen sevk tarihi
        baglan.ExecuteScript ("$('#*****').val(" & Cells(j, 8) & ").change()")         'Fiyat belirleme tarihi
        baglan.ExecuteScript ("$('#*****').val(" & Cells(j, 9) & ").change()")        'Belge tarihi
        baglan.Wait 4000
        baglan.FindElementByXPath("//*[@id='0']/div[3]/*****/span").Click 'ilk kalem için malzeme seçimi
        baglan.Wait 4000
        baglan.FindElementByXPath("//*[@id='0']/div[3]/*****/input").SendKeys Cells(j, 2).Value 'ilk kalem için malzeme seçimi
        baglan.Wait 4000
        baglan.FindElementByXPath("//*[@id='0']/div[3]/*****/a/span[1]").Click 'ilk kalem için malzeme seçimi
        baglan.Wait 4000
        baglan.ExecuteScript ("$('#*****').val(" & Cells(j, 3) & ").change()")      'Malzeme Miktari
        baglan.WaitForScript "!jQuery.active"
        baglan.Mouse.Click
        baglan.Wait 4000
        baglan.WaitForScript "!jQuery.active"
        baglan.ExecuteScript ("$('#*****').autoNumeric('set', " & Cells(j, 4).Value & ").change()") 'Birim fiyat
        baglan.WaitForScript "!jQuery.active"
        baglan.Wait 4000
        baglan.ExecuteScript ("$('select#*****').val(" & Cells(j, 10).Value & ").change()") 'depo seçimi *****:merkez depo
        baglan.WaitForScript "!jQuery.active"
        baglan.Wait 4000
        baglan.Mouse.Click
        nextt = 1
        baglan.Wait 2000
        For i = k To m
            If Cells(i, 1) <> Cells(i + 1, 1) Then
            baglan.Wait 2000
            Exit For
            Else
                'baglan.FindElementByXPath("//*[@id='*****']/a").Click 'malzeme eklemek için
                baglan.Wait 2000
                baglan.ExecuteScript ("$('#*****').click()")
                'baglan.FindElementById("*****").Click
                baglan.Wait 4000
                baglan.FindElementByXPath("//*[@id='" & nextt & "']/div/*****/span").Click 'diger kalem için malzeme seçimi
                baglan.Wait 4000
                baglan.FindElementByXPath("/html/body/div[" & nextt + 8 & "]/div/*****/input").SendKeys Cells(i + 1, 2).Value 'diger kalem için malzeme seçimi
                baglan.Wait 4000
                baglan.FindElementByXPath("//html/body/div[" & nextt + 8 & "]/div/*****/span[1]").Click 'diger kalem için malzeme seçimi
                baglan.Wait 4000
                baglan.ExecuteScript ("$('#" & nextt & "*****').val(" & Cells(i + 1, 3) & ").change()") 'Malzeme Miktari
                baglan.WaitForScript "!jQuery.active"
                baglan.Wait 4000
                baglan.Mouse.Click
                baglan.WaitForScript "!jQuery.active"
                baglan.ExecuteScript ("$('#" & nextt & "*****').autoNumeric('set', " & Cells(i + 1, 4).Value & ").change()") 'Birim fiyat
                baglan.WaitForScript "!jQuery.active"
                baglan.Wait 4000
                baglan.ExecuteScript ("$('select#" & nextt & "*****').val(" & Cells(i + 1, 10).Value & ").change()") 'depo seçimi *****:merkez depo
                baglan.WaitForScript "!jQuery.active"
                baglan.Wait 4000
                baglan.Mouse.Click
                baglan.Wait 4000
                nextt = nextt + 1
            End If
            h = h + 1
            k = k + 1
        Next i
        baglan.Wait 4000
        baglan.FindElementByXPath("//*[@id='*****']/div[2]/*****/div/button[1]").Click 'Kaydet butonu
        baglan.Wait 4000
    Exit For
    Next j
    k = k + 1
    h = h + 1
If Cells(h, 2) = "end" Then
Exit For
End If
Next r
End Sub

