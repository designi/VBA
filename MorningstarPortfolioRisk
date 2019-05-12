Private Sub CommandButton1_Click()
 
 ' Vol measures 1 Ticker at a time
 
 Dim ie As Object, i As Long, doc As Object, wb As Excel.Workbook, ws As Excel.Worksheet
   Dim hTable As Object, tbody As Object, strText As String
   Dim hBody As Object, hTR As Object, hTD As Object
   Dim tb As Object, bb As Object, Tr As Object, Td As Object, val1 As Object ', val1 As Object
   Dim y As Long, z As Long, val0 As Object, f As Long ', j As Long
   Application.ScreenUpdating = 0
     Sheets(1).Activate
     Sheets(1).Range("A2:F7").ClearContents ' Clear Box
     Application.StatusBar = "Do Not Touch"
     
     Set wb = Excel.ActiveWorkbook
     Set ws = wb.ActiveSheet
     Set ie = CreateObject("InternetExplorer.Application")
     ie.Visible = False 'Turn on internet explorer view

     y = 1
     z = 1
     i = 2
     f = 1 ' set row label loop iteration

    
        ie.navigate "http://performance.morningstar.com/funds/etf/ratings-risk.action?t=" & Sheets(1).Cells(3, 9) & "&region=usa&culture=en-US"
        On Error Resume Next
        Do While ie.Busy: DoEvents: Loop
        Do While ie.ReadyState <> 4: DoEvents: Loop
        
        Application.Wait (Now + TimeValue("00:00:02"))
        
        Set doc = ie.document
          Set hTable = doc.getelementbyid("div_volatility").getElementsByClassName("r_table2 text2")
            For Each tb In hTable
            '''''
             Set hBody = tb.getElementsByTagName("tbody")
                 For Each bb In hBody
                     Set hTR = bb.getElementsByTagName("tr")
                         For Each Tr In hTR
                              Set hTD = Tr.getElementsByTagName("td")
                                                y = 1 ' Resets back to column A
                                                j = 1
                                                For Each Td In hTD
                                                                                                        
                                                         Sheets(1).Cells(z + 1, y + 1).Value = Td.innerText ' Inner table data '  core content and statistics
                                                         If y + 11 <= 14 Then
                                                         Set val0 = doc.getElementsByClassName("row_lbl")(y + 11) 'Row Labels ---> Best-Fit Index, Category, etc....
                                                         If y = 1 Then
                                                         Sheets(1).Cells(f + y + 1, 1).Value = val0.innerText
                                                         ElseIf y = 2 Then Sheets(1).Cells(f + y + j + 1, 1).Value = val0.innerText
                                                         Else: If y = 3 Then Sheets(1).Cells(f + y + j + 1 + 1, 1).Value = val0.innerText
                                                         End If
                                                         End If

                                                y = y + 1
                                                
                                                Next Td
                                                 
                                              DoEvents
                                              
                                           z = z + 1
                                           
                                Next Tr
                            Exit For
                        Next bb
                    Exit For
                Next tb
                

        '''''
    Range("i3").Value = UCase(ActiveSheet.Range("i3"))
    Application.Wait (Now + TimeValue("00:00:02"))
    Application.StatusBar = "Ready"
    Set ie = Nothing
    ie.Quit
Application.ScreenUpdating = 1
End Sub



