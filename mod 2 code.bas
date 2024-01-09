VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock():
    
    ' declare all variables
    Dim ticker As String
    Dim YearEnd As Double
    Dim tracker As Integer
    Dim YearStart As Double
    Dim volume As Double
    Dim Change As Double
    Dim ws As Worksheet
    
    ' make this cycle through each worksheet
    For Each ws In Worksheets
    
    ' set variable values
    volume = 0
    tracker = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    YearStart = ws.Range("C2").Value
    
    ' add column names to worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"


   ' loop through all lines of data
    For i = 2 To LastRow
    
    
        ' determine if we are within the same ticker info, this code is from the credit card activity in class
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            
            ' set the ticker name
            ticker = ws.Cells(i, 1).Value
            
            ' set the outputs wanted for the summary table
            YearEnd = ws.Cells(i, 6).Value
            Change = YearEnd - YearStart
            volume = volume + ws.Cells(i, 7).Value
            
            ' add outputs to the summary table
            ws.Range("I" & tracker).Value = ticker
            ws.Range("J" & tracker).Value = Change
            ws.Range("K" & tracker).Value = (Change / YearStart)
            ws.Range("L" & tracker).Value = volume
            
            
            ' this code came from https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
            ws.Range("K" & tracker).NumberFormat = "0.00%"
            
            ' add color to cells to yearly change using conditional formatting, code found at https://www.automateexcel.com/vba/conditional-formatting/
            Dim MyRange As Range
            Set MyRange = ws.Range("J2:J3001")
            MyRange.FormatConditions.Delete
            MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
            MyRange.FormatConditions(1).Interior.ColorIndex = 3
            MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
            MyRange.FormatConditions(2).Interior.ColorIndex = 4
            
            ' do the same for the percentage change, code source same as above
            Dim PerRange As Range
            Set PerRange = ws.Range("K2:K3001")
            PerRange.FormatConditions.Delete
            PerRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
            PerRange.FormatConditions(1).Interior.ColorIndex = 3
            PerRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
            PerRange.FormatConditions(2).Interior.ColorIndex = 4
            
            ' reset variables
            tracker = tracker + 1
            start = i + 1
            volume = 0
            YearStart = ws.Cells(i + 1, 3).Value
        
        Else
        
            ' add to stock volume if within the same ticker
            volume = volume + ws.Cells(i, 7).Value
    
    End If

Next i

    
    ' set  variable to find greatest increase
    Dim HighestChange As Double
    HighestChange = 0
    
    For j = 2 To 3001
        
        ' run loop to find new highest percentage increase and add values to table
        If ws.Cells(j, 11).Value > HighestChange Then
        
            HighestChange = ws.Cells(j, 11).Value
            ws.Cells(2, 17).Value = HighestChange
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value

        End If

    Next j
    
   ' set variable to find greatest decrease
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    
    ' run for loop to find minimum
    For k = 2 To 3001
        
        'replace variable with new lower number and add values to table
        If ws.Cells(k, 11).Value < GreatestDecrease Then
        
            GreatestDecrease = ws.Cells(k, 11).Value
            ws.Cells(3, 17).Value = GreatestDecrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
        
        End If
    
    Next k
    
    
    ' set  variable to find greatest stock volume increase
    Dim HighestStock As LongLong
    HighestStock = 0
    
    For l = 2 To 3001
        
        ' run loop to find highest stock increase increase and add values to table
        If ws.Cells(l, 12).Value > HighestStock Then
        
            HighestStock = ws.Cells(l, 12).Value
            ws.Cells(4, 17).Value = HighestStock
            ws.Cells(4, 16).Value = ws.Cells(l, 9).Value

        End If

    Next l
    
Next ws


End Sub




