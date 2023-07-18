Attribute VB_Name = "Module1"
Sub YearlyStocks()


Dim ws As Variant


'--------------------------------------------------------------------------------------------------------
'                                       ITERATION FOR EACH WORKSHEETS
'--------------------------------------------------------------------------------------------------------
For Each ws In Worksheets
'---------------------------...---------------------------------------...--------------------------------
'                                       DEFINING VARIABLES
'--------------------------------------------------------------------------------------------------------

Dim WSN As String

Dim YearlyChange As Long
Dim PerChange As Double
Dim TotalSChange As Long

Dim GI As Double
Dim GD As Double
Dim GV As Double

Dim LastRow As Long

Dim i As Long
Dim x As Long
Dim Count As Long

x = 2
Count = 2

WSN = ws.Name
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Range("A1:L1").Interior.Color = RGB(0, 0, 0)
Range("A1:L1").Font.Color = RGB(255, 255, 255)
Worksheets(WSN).Columns("A:Z").AutoFit

'.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.
'.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.
'.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.

For i = 2 To LastRow:




'--------------------------------------------------------------------------------------------------------
'                                       THE TICKER SYMBOL
'--------------------------------------------------------------------------------------------------------

ws.Cells(1, 9) = "Ticker"

ws.Cells(Count, 9) = ws.Cells(i, 1).Value



'--------------------------------------------------------------------------------------------------------
'                                       YEARLY CHANGE
'--------------------------------------------------------------------------------------------------------

ws.Cells(1, 10) = "Yearly Change"

ws.Cells(Count, 10).Value = ws.Cells(i, 6).Value - ws.Cells(x, 3).Value

    
    
        If ws.Cells(Count, 10).Value < 0 Then
        
            ws.Cells(Count, 10).Interior.Color = RGB(255, 0, 0)
        Else
            ws.Cells(Count, 10).Interior.Color = RGB(0, 255, 0)
        End If

    
'--------------------------------------------------------------------------------------------------------
'                                       PERCENT CHANGE
'--------------------------------------------------------------------------------------------------------

ws.Cells(1, 11) = "Percent Change"



    If ws.Cells(x, 3).Value <> 0 Then
    
        PerChange = ((ws.Cells(i, 6).Value - ws.Cells(x, 3).Value) / ws.Cells(x, 3).Value)
    
        ws.Cells(Count, 11).Value = Format(PerChange, "Percent")
    
    End If
    


'--------------------------------------------------------------------------------------------------------
'                                       TOTAL VOLUME
'--------------------------------------------------------------------------------------------------------

ws.Cells(1, 12) = " Total Stock volume "

        
ws.Cells(Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(x, 7), ws.Cells(i, 7)))


        
        
'--------------------------------------------------------------------------------------------------------




Count = Count + 1
x = i + 1


Next i

'.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.
'.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.
'.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.-.









'--------------------------------------------------------------------------------------------------------
'                                       GREATEST - TABLE
'--------------------------------------------------------------------------------------------------------


ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
Range("P1:Q1").Interior.Color = RGB(0, 0, 0)
Range("P1:Q1").Font.Color = RGB(255, 255, 255)

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
Range("O2:O4").Interior.Color = RGB(0, 0, 0)
Range("O2:O4").Font.Color = RGB(255, 255, 255)



For i = 2 To LastRow



GI = ws.Cells(2, 11).Value

    If ws.Cells(i, 11).Value > GI Then
        GI = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    Else
        GI = GI
    End If
    


GD = ws.Cells(2, 11).Value

    If ws.Cells(i, 11).Value < GD Then
        GD = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    Else
        GD = GD
    End If
        
    


GV = ws.Cells(2, 12).Value


    If ws.Cells(i, 12).Value > GV Then
        GV = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    Else
        GV = GV
    End If
    
    
ws.Cells(2, 17).Value = Format(GI, "Percent")
ws.Cells(3, 17).Value = Format(GD, "Percent")
ws.Cells(4, 17).Value = Format(GV, "Scientific")


Next i
'--------------------------------------------------------------------------------------------------------




Next ws
End Sub




