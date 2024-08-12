Sub WhatsAppWeb()

    Dim bot As New WebDriver
    Dim ks As New Keys
    Dim attempts As Integer
    Dim maxAttempts As Integer
    
    bot.Start "chrome", "https://web.whatsapp.com/"
    bot.Get "/"
    
    MsgBox "Please scan the QR code. After you are logged in, please confirm this message box by clicking OK"
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
        searchtext = Sheets(1).Range("A" & i).Value
        textmessage = Sheets(1).Range("B" & i).Value
        
        ' Retry mechanism for StaleElementReferenceError
        attempts = 0
        maxAttempts = 3
        Do
            On Error Resume Next
            bot.FindElementByXPath("//*[@id='side']/div[1]/div/div[2]/div[2]/div/div/p").Click
            On Error GoTo 0
            bot.FindElementByXPath("//*[@id='side']/div[1]/div/div[2]/div[2]/div/div/p").Click
            
            bot.Wait (1000)
            
            bot.SendKeys (searchtext)
            
            bot.Wait (1000)
            
            bot.SendKeys (ks.Enter)
            
            bot.Wait (1000)
            
            bot.SendKeys (textmessage)
            
            bot.Wait (1000)
            
            On Error Resume Next
            bot.SendKeys (ks.Enter)
            On Error GoTo 0
            
            ' Check for StaleElementReferenceError and retry if necessary
            If Err.Number = 0 Then
                Exit Do
            Else
                Err.Clear
                attempts = attempts + 1
                If attempts >= maxAttempts Then
                    ' Optionally, handle the error or log it if needed
                    MsgBox "Error: StaleElementReferenceError"
                    Exit For
                End If
                ' Optionally, wait a short period before retrying
                bot.Wait (1000)
            End If
        Loop
        
    Next i
    
    MsgBox "Done"
    
End Sub