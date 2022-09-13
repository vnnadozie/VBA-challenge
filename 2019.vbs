VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Module2()


    Dim Ticker As String

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"

    Dim Vol_Total As Double
    Vol_Total = 0

    Dim Sum_Rows As Integer
    Sum_Rows = 2

    For i = 2 To 753001
        
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            
            Ticker = Cells(i, 1).Value
            Vol_Total = Vol_Total + Cells(i, 7).Value
            
            Range("I" & Sum_Rows) = Ticker
            Range("L" & Sum_Rows) = Vol_Total
            
            Sum_Rows = Sum_Rows + 1
            
            Vol_Total = 0
            
        Else
        
            Vol_Total = Vol_Total + Cells(i, 7).Value
    End If
 
Next i


    Dim TickStartRow As Long
    Dim TickEndRow As Long
    Dim Sum_Row As Integer
    Dim Percentage_Change As Long
    
    Sum_Row = 2
    Last_Row = Cells(Rows.Count, 9).End(xlUp).Row

    
    For i = 2 To Last_Row
        
        TickStartRow = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1, 1), LookAt:=xlWhole).Row
        TickEndRow = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1, 1), LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
        
        Range("J" & Sum_Row).Value = Range("F" & TickEndRow).Value - Range("C" & TickStartRow).Value
        
        Range("K" & Sum_Row).Value = (Range("F" & TickEndRow).Value - Range("C" & TickStartRow).Value) / Range("C" & TickStartRow).Value
        Range("K" & Sum_Row).NumberFormat = "0.00%"

        Sum_Row = Sum_Row + 1
        
    Next i

    
    Dim Color_Range As Range
    
    Set Color_Range = Worksheets("2019").Range("J2:J91")
    
    For Each Cell In Color_Range
        
        If Cell.Value < 0 Then
        Cell.Interior.ColorIndex = 3 'Red
        
        Else
        Cell.Interior.ColorIndex = 4 'Green
        
        End If
    Next
        
End Sub

Sub Bonus()

'I can't figure out how to print the Ticker

    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Dim Percentage As Range
    Dim Greatest_Decrease As Range
    Dim Greatest_Volume As Range
    
    Set Percentage = Worksheets("2019").Range("K2:K91")
    Set Volume = Worksheets("2019").Range("L2:L91")
    
    decrease = Application.WorksheetFunction.Min(Percentage)
    increase = Application.WorksheetFunction.Max(Percentage)
    vol = Application.WorksheetFunction.Max(Volume)
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    Cells(2, 17).Value = increase
    Cells(3, 17).Value = decrease
    Cells(4, 17).Value = vol
End Sub

