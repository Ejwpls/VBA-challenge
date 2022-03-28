Attribute VB_Name = "Module1"
Sub getThings()

Dim rng As Range
Dim wb As Workbook

Dim ticker_add As String

    Set ws = ActiveWorkbook.ActiveSheet
    Set rng = ws.UsedRange
    
    'Adress name for data name column
    ticker_add = rng.Columns(1).Address
    
    Range("J:J").Clear
    
    'Get all ticker unique name
    rng.Range(ticker_add).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("J1"), Unique:=True
    

Dim ticker_i As Double
Dim ticker_arr()

    Dim open_price() As Double
    Dim close_price() As Double

    'Get nummber of ticker
    Dim ticker_rng As Range
    Set ticker_rng = ws.Range("J:J")
    
    Dim n As Integer
    
    ticker_i = 0
    
    For n = 0 To 10000
        If IsEmpty(ticker_rng.Cells(n + 1, 1)) = False Then
            ticker_i = ticker_i + 1
        Else
            Exit For
        End If
        
    Next n

    'Remove header
    ticker_i = ticker_i - 2
    
    ReDim ticker_arr(ticker_i)
    
    Dim ticker As Variant
    
    For n = 0 To ticker_i
        ticker_arr(n) = ticker_rng.Cells(n + 2, 1).Value2
    Next n
    
    'Rename all the header
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    ReDim open_price(ticker_i)
    ReDim close_price(ticker_i)
    
    Dim startdate As Double
    Dim enddate As Double
    
    Dim stguid As String
    Dim endguid As String
    Dim a As String
    
    For n = 0 To ticker_i
        startdate = WorksheetFunction.MinIfs(Range("B:B"), Range("A:A"), "=" & ticker_arr(n))
        enddate = WorksheetFunction.MaxIfs(Range("B:B"), Range("A:A"), "=" & ticker_arr(n))
        
        open_price(n) = WorksheetFunction.MaxIfs(Range("C:C"), Range("A:A"), "=" & ticker_arr(n), Range("b:b"), "=" & startdate)
        close_price(n) = WorksheetFunction.MaxIfs(Range("F:F"), Range("A:A"), "=" & ticker_arr(n), Range("b:b"), "=" & enddate)
        
        ws.Cells(n + 2, 11).Value = close_price(n) - open_price(n)
        If open_price(n) = 0 Then
            ws.Cells(n + 2, 12) = 0
        Else
            ws.Cells(n + 2, 12).Value = (close_price(n) - open_price(n)) / open_price(n) * 100
        End If
        
            If ws.Cells(n + 2, 11).Value < 0 Then
                    With ws.Cells(n + 2, 12).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                Else
                    With ws.Cells(n + 2, 11).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 5287936
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
            End If
                
        ws.Cells(n + 2, 13).Value = WorksheetFunction.SumIfs(Range("G:G"), Range("A:A"), "=" & ticker_arr(n))
        
    Next n
        
    'Challenge
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(2, 17) = WorksheetFunction.Max(Range("L:L"))
    ws.Cells(2, 16) = WorksheetFunction.XLookup(ws.Cells(2, 17), Range("L:L"), Range("J:J"))
    
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(3, 17) = WorksheetFunction.Min(Range("L:L"))
    ws.Cells(3, 16) = WorksheetFunction.XLookup(ws.Cells(3, 17), Range("L:L"), Range("J:J"))
    
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(4, 17) = WorksheetFunction.Max(Range("M:M"))
    ws.Cells(4, 16) = WorksheetFunction.XLookup(ws.Cells(4, 17), Range("m:m"), Range("J:J"))
    
End Sub


Sub Runeveryworksheet():
    
    Dim PauseTime, Start, Finish, TotalTime
    
    Application.ScreenUpdating = False
    
    Start = Timer    ' Set start time.
    
    Dim countWS As Integer
    Dim i As Integer
    Dim ws As Variant
    
    countWS = ThisWorkbook.Worksheets.Count
    
    For i = 1 To countWS
        Sheets(i).Activate
        Call getThings
    Next i
    
    Finish = Timer    ' Set end time.
    TotalTime = Finish - Start    ' Calculate total time.
    MsgBox "Paused for " & TotalTime & " seconds"
    
    Application.ScreenUpdating = True

End Sub

