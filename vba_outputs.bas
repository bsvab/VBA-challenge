Attribute VB_Name = "Outputs"
Sub PopulateAll()
    
'    '////////// following section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
'    Dim StartTime As Double
'    Dim SecondsElapsed As Double
'    Dim MinutesElapsed As Double
'    Dim HoursElapsed As Double
'    StartTime = Timer
'    '\\\\\\\\\\ previous section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
    
    Application.DisplayStatusBar = True
    Application.StatusBar = "Macro is running...... patience please :)"
    DoEvents
    
    Call ClearOutputs
    Call FirstOutputs
    Call SecondOutputs
    DoEvents
    
    Application.StatusBar = False
    
'    MsgBox ("DONE! " & Now())
    
'    '////////// following section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
'    SecondsElapsed = Round(Timer - StartTime, 2)
'    MinutesElapsed = Round(SecondsElapsed / 60, 2)
'    HoursElapsed = Round(MinutesElapsed / 60, 2)
'    Debug.Print "Run Time: " & SecondsElapsed & " sec / " & MinutesElapsed & " min / " & HoursElapsed & " hrs"
'    MsgBox "Run Time: " & SecondsElapsed & " sec / " & MinutesElapsed & " min / " & HoursElapsed & " hrs", vbInformation
'    '\\\\\\\\\\ previous section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time

End Sub

Sub FirstOutputs()
    
    '////////// following section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
    'Dim StartTime As Double
    'Dim SecondsElapsed As Double
    'Dim MinutesElapsed As Double
    'Dim HoursElapsed As Double
    'StartTime = Timer
    '\\\\\\\\\\ previous section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
    
    For Each ws In Worksheets
        
        'Setting output headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        DoEvents
        
        Dim dataLR As Long
        dataLR = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        'Start iterating to fill in data in output columns...
        
        Dim outputLR As Long
        outputLR = ws.Cells(Rows.Count, "I").End(xlUp).row
        
        Dim ticker As String
        Dim firstopen As Double
        Dim lastclose As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        
        Dim volume As Single
        volume = 0
        
        Dim tickertracker As Long
        tickertracker = 0
        
        Dim outputrowtracker As Long
        outputrowtracker = 2
    
        Dim firstdate As Long
        
        If ws.Name = "2018" Then
            firstdate = 20180102
        ElseIf ws.Name = "2019" Then
            firstdate = 20190102
        ElseIf ws.Name = "2020" Then
            firstdate = 20200102
        End If
        
        Dim lastdate As Long
        
        If ws.Name = "2018" Then
            lastdate = 20181231
        ElseIf ws.Name = "2019" Then
            lastdate = 20191231
        ElseIf ws.Name = "2020" Then
            lastdate = 20201231
        End If
        
        For i = 2 To dataLR
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                volume = volume + ws.Cells(i, 7).Value
'                Debug.Print volume
                ws.Range("L" & outputrowtracker).Value = volume 'Populating "Total Stock Volume" column
                
                If ws.Cells(i, 2).Value = firstdate Then
                    firstopen = ws.Cells(i, 3).Value
'                    Debug.Print firstopen

                    ElseIf ws.Cells(i, 2).Value = lastdate Then
                        lastclose = ws.Cells(i, 6).Value
'                        Debug.Print lastclose
                        yearlychange = lastclose - firstopen
'                        Debug.Print yearlychange
                        ws.Range("J" & outputrowtracker).Value = yearlychange 'Populating "Yearly Change" column with the delta of the end of year close and start of year open

                        If yearlychange > 0 Then
                            ws.Range("J" & outputrowtracker).Interior.ColorIndex = 4 'Color positive "Yearly Change" values as green
                        ElseIf yearlychange < 0 Then
                            ws.Range("J" & outputrowtracker).Interior.ColorIndex = 3 'Color negative "Yearly Change" values as red
                        End If

                        percentchange = yearlychange / firstopen
                        ws.Range("K" & outputrowtracker).Value = percentchange 'populating "Percent Change" column with the % of the delta of the end of year close and start of year open vs the start of year open
                        ws.Range("K" & outputrowtracker).NumberFormat = "0.00%" 'fix format of percentchange to be a percent

                        If percentchange > 0 Then
                            ws.Range("K" & outputrowtracker).Interior.ColorIndex = 4 'Color positive "Percent Change" values as green
                        ElseIf percentchange < 0 Then
                            ws.Range("K" & outputrowtracker).Interior.ColorIndex = 3 'Color negative "Percent Change" values as red
                        End If

                End If
                
                outputrowtracker = outputrowtracker + 1 'Increaseing row tracker by 1 to move on to next ticker
                tickertracker = tickertracker + 1 'Increasing ticker tracker by 1 to move on to next ticker
                volume = 0 'Reset volume to zero before iterating next ticker
'                Debug.Print volume
                
                Else
                    ticker = ws.Cells(i, 1).Value
'                    Debug.Print ticker
                    ws.Range("I" & outputrowtracker).Value = ticker 'Populating "Ticker" column with the ticker letters
                    
'                    firstdate = Application.WorksheetFunction.MinIfs(data.Columns(2), data.Columns(1), ticker)
'                    Debug.Print firstdate
'
'                    lastdate = Application.WorksheetFunction.MaxIfs(data.Columns(2), data.Columns(1), ticker)
'                    Debug.Print lastdate
                    
                    If ws.Cells(i, 2).Value = firstdate Then
                        firstopen = ws.Cells(i, 3).Value
'                        Debug.Print firstopen

                        ElseIf ws.Cells(i, 2).Value = lastdate Then
                            lastclose = ws.Cells(i, 6).Value
'                            Debug.Print lastclose
                            yearlychange = lastclose - firstopen
'                            Debug.Print yearlychange
                            ws.Range("J" & outputrowtracker).Value = yearlychange 'Populating "Yearly Change" column with the delta of the end of year close and start of year open

                            If yearlychange > 0 Then
                                ws.Range("J" & outputrowtracker).Interior.ColorIndex = 4 'Color positive "Yearly Change" values as green and negative values as red
                            ElseIf yearlychange < 0 Then
                                ws.Range("J" & outputrowtracker).Interior.ColorIndex = 3 'Color positive "Yearly Change" values as green and negative values as red
                            End If

                            percentchange = yearlychange / firstopen
                            ws.Range("K" & outputrowtracker).Value = percentchange 'populating "Percent Change" column with the % of the delta of the end of year close and start of year open vs the start of year open
                            ws.Range("K" & outputrowtracker).NumberFormat = "0.00%" 'fix format of percentchange to be a percent

                            If percentchange > 0 Then
                                ws.Range("K" & outputrowtracker).Interior.ColorIndex = 4 'Color positive "Percent Change" values as green
                            ElseIf percentchange < 0 Then
                                ws.Range("K" & outputrowtracker).Interior.ColorIndex = 3 'Color negative "Percent Change" values as red
                            End If

                    End If
                    volume = volume + ws.Cells(i, 7).Value
'                    Debug.Print volume
            End If
            
        Next i
        
        ws.Range("R22").Value = "NOTE: No instructions were given regarding 0 yearly/percent change values, so no conditional format was applied when = 0"
        ws.Range("R22").Font.ColorIndex = 3
        
        ws.Columns("I:L").AutoFit
        
        DoEvents
        
    Next ws
    
    '////////// following section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
    'SecondsElapsed = Round(Timer - StartTime, 2)
    'MinutesElapsed = Round(SecondsElapsed / 60, 2)
    'HoursElapsed = Round(MinutesElapsed / 60, 2)
    'Debug.Print "Run Time: " & SecondsElapsed & " sec / " & MinutesElapsed & " min / " & HoursElapsed & " hrs"
    'MsgBox "Run Time: " & SecondsElapsed & " sec / " & MinutesElapsed & " min / " & HoursElapsed & " hrs", vbInformation
    '\\\\\\\\\\ previous section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time

End Sub


Sub SecondOutputs()

    '////////// following section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
    'Dim StartTime As Double
    'Dim SecondsElapsed As Double
    'Dim MinutesElapsed As Double
    'Dim HoursElapsed As Double
    'StartTime = Timer
    '\\\\\\\\\\ previous section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time

    For Each ws In Worksheets

        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"

        Dim outputLR As Long
        outputLR = ws.Cells(Rows.Count, "I").End(xlUp).row

        Dim greatestincrease As Double
        Dim GItickerrow As Long
        Dim GIticker As String
        Dim greatestdecrease As Double
        Dim GDtickerrow As Long
        Dim GDticker As String
        Dim greatestvolume As Single
        Dim GVtickerrow As Long
        Dim GVticker As String

        greatestincrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & outputLR))
        'Debug.Print greatestincrease
        ws.Range("P2").Value = greatestincrease
        ws.Range("P2").NumberFormat = "0.00%"      'fix format to be a percent

        GItickerrow = Application.WorksheetFunction.Match(greatestincrease, ws.Range("K2:K" & outputLR), 0) + 1
        GIticker = ws.Cells(GItickerrow, 9).Value
        ws.Range("O2").Value = GIticker

        greatestdecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & outputLR))
        'Debug.Print Application.WorksheetFunction.Min(ws.Range("K2:K" & outputLR))
        ws.Range("P3").Value = greatestdecrease
        ws.Range("P3").NumberFormat = "0.00%"      'fix format to be a percent

        GDtickerrow = Application.WorksheetFunction.Match(greatestdecrease, ws.Range("K2:K" & outputLR), 0) + 1
        GDticker = ws.Cells(GDtickerrow, 9).Value
        ws.Range("O3").Value = GDticker

        greatestvolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & outputLR))
        'Debug.Print Application.WorksheetFunction.Max(ws.Range("L2:L" & outputLR))
        ws.Range("P4").Value = greatestvolume

        GVtickerrow = Application.WorksheetFunction.Match(greatestvolume, ws.Range("L2:L" & outputLR), 0) + 1
        GVticker = ws.Cells(GVtickerrow, 9).Value
        ws.Range("O4").Value = GVticker

        ws.Columns("N:P").AutoFit

        DoEvents

    Next ws

    '////////// following section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
    'SecondsElapsed = Round(Timer - StartTime, 2)
    'MinutesElapsed = Round(SecondsElapsed / 60, 2)
    'HoursElapsed = Round(MinutesElapsed / 60, 2)
    'Debug.Print "Run Time: " & SecondsElapsed & " sec / " & MinutesElapsed & " min / " & HoursElapsed & " hrs"
    'MsgBox "Run Time: " & SecondsElapsed & " sec / " & MinutesElapsed & " min / " & HoursElapsed & " hrs", vbInformation
    '\\\\\\\\\\ previous section of code sourced from (https://www.thespreadsheetguru.com/vba-calculate-macro-run-time/) to report macro run time
    
End Sub

Sub ClearOutputs()
    
    For Each ws In Worksheets
        
        ws.Range("I:P").ClearContents
        ws.Range("I:P").ClearFormats
        ws.Range("I:P").ColumnWidth = 8
        ws.Range("R22").ClearContents
        ws.Range("R22").ClearFormats
        
    Next ws
    
End Sub



