VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Module2()


    'Set initial variable for holding the ticker
    Dim Ticker As String

    'Create new column headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Set an initial volume for the volume total
    Dim Vol_Total As Double
    Vol_Total = 0

    'Keep track of summary rows
    Dim Sum_Rows As Integer
    Sum_Rows = 2

    'Loop through all Tickers
    For i = 2 To 22771
        
        'Check if the Ticker is the same, if it is not...
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            
            'Set the Ticker
            Ticker = Cells(i, 1).Value
            'Add to the Volume Total
            Vol_Total = Vol_Total + Cells(i, 7).Value
            
            'Print the Ticker and Vol in the summary rows
            Range("I" & Sum_Rows) = Ticker
            Range("L" & Sum_Rows) = Vol_Total
            
            'Add one to the summary row for the next Ticker
            Sum_Rows = Sum_Rows + 1
            
            'Rest Vol Value
            Vol_Total = 0
            
        'If the cell immediately following is the same Ticker
        Else
        
            'Add to the Ticker Total
            Vol_Total = Vol_Total + Cells(i, 7).Value
    End If
 
Next i


    'Find start and end rows with unique Ticker
    Dim TickStartRow As Long
    Dim TickEndRow As Long
    Dim Sum_Row As Integer
    Dim Percentage_Change As Long
    
    Sum_Row = 2
    LastRow1 = Cells(Rows.Count, 9).End(xlUp).Row

    
    'Loop through all tickers
    For i = 2 To LastRow1
        
        'Find start and end rows
        TickStartRow = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1, 1), LookAt:=xlWhole).Row
        TickEndRow = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1, 1), LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
        
        'Print Year Change in summary row
        Range("J" & Sum_Row).Value = Range("F" & TickEndRow).Value - Range("C" & TickStartRow).Value
        
        'Print Percentage Change in summary row
        Range("K" & Sum_Row).Value = (Range("F" & TickEndRow).Value - Range("C" & TickStartRow).Value) / Range("C" & TickStartRow).Value
        Range("K" & Sum_Row).NumberFormat = "0.00%"
        'Add one to the summary row for the next ticker
        Sum_Row = Sum_Row + 1
        
    Next i

    
   'Make a range variable
    Dim Color_Range As Range
    
    'Set range in column J
    Set Color_Range = Worksheets("2018").Range("J2:J91")
    
    'Make loop for the Yearly Change values
    For Each Cell In Color_Range
        
        'If change is negative then...
        If Cell.Value < 0 Then
        Cell.Interior.ColorIndex = 3 'Red
        
        'If change is positive then...
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
    
    Set Percentage = Worksheets("2018").Range("K2:K91")
    Set Volume = Worksheets("2018").Range("L2:L91")
    
    decrease = Application.WorksheetFunction.Min(Percentage)
    increase = Application.WorksheetFunction.Max(Percentage)
    vol = Application.WorksheetFunction.Max(Volume)
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    Cells(2, 17).Value = increase
    Cells(3, 17).Value = decrease
    Cells(4, 17).Value = vol
End Sub

