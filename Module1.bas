Attribute VB_Name = "Module1"
Sub StockAnalyzer()
'variable creation
Dim TickCount As Double 'hold count for tickers
Dim TickVolume As Double 'volume of current ticker
Dim TickOpen As Double 'remember opening value
Dim TickClose As Double 'remember closing value
Dim OutCount As Double 'counter for data output
Dim CurrentTick As String 'hold currently worked on ticker
Dim LastTick As String 'hold last tick
Dim GreatestIncrease As Double ' hold greatest increase
Dim GreatestDecrease As Double 'hold greatest decrease
Dim GreatestVolume As Double 'hold greatest volume
Dim rowopen As Double 'stores the open per day
Dim Sheets As Integer 'variable for sheets in workbook code from microsoft support (https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0)




'set up table for output
'Cells(1, 9).Value = "Ticker"
'Cells(1, 10).Value = "Quarterly Change"
'Cells(1, 11).Value = "Percent Change"
'Cells(1, 12).Value = "Total Stock Volume"
'Cells(1, 16).Value = "Ticker"
'Cells(1, 17).Value = "Value"
'Cells(2, 15).Value = "Greatest % Increase"
'Cells(3, 15).Value = "Greatest % Decrease"
'Cells(4, 15).Value = "Greatest Total Volume"

'OutCount = 2


'Prime three loops
'CurrentTick = Cells(2, 1).Value
'LastTick = Cells(2, 1).Value
' TickOpen = Cells(2, 3).Value
'TickCount = 2
Sheets = ActiveWorkbook.Worksheets.Count ' get count of sheets in workbook code from microsoft support (https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0)

For i = 1 To Sheets 'loop through sheets, code from microsoft support (https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0)
Worksheets(i).Activate

'set up table for output
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quarterly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"



'sanitize values for each worksheet
CurrentTick = Cells(2, 1).Value
LastTick = Cells(2, 1).Value
 TickOpen = Cells(2, 3).Value
TickCount = 2
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0
OutCount = 2


    'loop through data
    Do While Not (CurrentTick = "")
        'mini loop for same stock
        Do While LastTick = CurrentTick
            TickVolume = TickVolume + Cells(TickCount, 7).Value
            rowopen = Cells(TickCount, 3).Value
            TickClose = Cells(TickCount, 6).Value
            TickCount = TickCount + 1
            CurrentTick = Cells(TickCount, 1).Value
        Loop
        
        'Create analysis table
        Cells(OutCount, 9).Value = LastTick 'Tick name
        
        Cells(OutCount, 10).Value = TickClose - TickOpen 'Quarterly Change
        'Color Cells by value of change
        If Cells(OutCount, 10).Value > 0 Then 'Positive
            Cells(OutCount, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf Cells(OutCount, 10).Value < 0 Then 'Negative
            Cells(OutCount, 10).Interior.Color = RGB(255, 0, 0)
        End If
        Cells(OutCount, 11).Value = FormatPercent((TickClose - TickOpen) / TickOpen) 'Percentage of change
        Cells(OutCount, 12).Value = TickVolume 'Volume Moved
        
        'Check for greatness
        If TickVolume > GreatestVolume Then 'greatest volume
            GreatestVolume = TickVolume
            Cells(4, 16).Value = LastTick
            Cells(4, 17).Value = GreatestVolume
        End If
        If Cells(OutCount, 10).Value / TickOpen > GreatestIncrease Then 'greatest increase
            GreatestIncrease = Cells(OutCount, 10).Value / TickOpen
            Cells(2, 16).Value = LastTick
            Cells(2, 17).Value = FormatPercent(GreatestIncrease)
        End If
        If Cells(OutCount, 10).Value / TickOpen < GreatestDecrease Then 'greatest decrease
            GreatestDecrease = Cells(OutCount, 10).Value / TickOpen
            Cells(3, 16).Value = LastTick
            Cells(3, 17).Value = FormatPercent(GreatestDecrease)
        End If
        OutCount = OutCount + 1
        
        'prime next loop
        TickVolume = Cells(TickCount, 7).Value
        TickOpen = Cells(TickCount, 3).Value
        LastTick = CurrentTick
        TickCount = TickCount + 1
        CurrentTick = Cells(TickCount, 1).Value
    
    Loop
    

Next i ' move to next worksheet in workbook code from microsoft support (https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0)

End Sub
 4 8 0 9 6 5 , 2 0 4 8 0 9 8 5 , 2 0 4 8 0 9 7 0 , 2 1 3 4 6 5 0 8 , 2 2 3 4 9 1 8 6 , 2 2 3 4 9 1 2 6 , 2 3 6 7 1 6 8 5 , 3 7 3 0 7 0 2 5 , 2 2 3 6 4 6 7 6 , 2 2 3 6 0 4 0 3 , 2 3 4 5 3 7 2 8 , 2 2 3 6 0 4 0 9 , 3 6 4 6 0 3 1 2 , 1 8 9 9 0 0 2 0 , 2 5 2 4 9 0 3 8 , 1 7 6 9 5 6 2 0 , 3 6 5 0 1 1 9 7 , 6 3 6 6 2 8 8 , 5 8 7 0 3 0 4 , 3 7 0 7 0 4 3 5 , 3 7 0 7 0 4 6 5 , 2 0 5 5 8 2 7 2 , 2 1 0 3 4 7 5 8 , 3 7 3 6 5 0 5 8 , 3 8 0 1 3 0 7 7 , 3 8 3 3 6 4 8 3 , 3 8 5 5 9 8 0 9 , 3 8 9 1 6 3 5 6 , 5 2 7 2 5 9 3 ,