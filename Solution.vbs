VBA Homework

Sub Moving_Sheets()

' Define Variables
Dim mainworkBook As Workbook
Dim namesheet As String

Set mainworkBook = ActiveWorkbook

For i = 1 To mainworkBook.Sheets.Count


namesheet = mainworkBook.Sheets(i).Name


'moving to next sheet
Sheets(namesheet).Select

'Below this line add routines

OrderByDate
clear_data
Get_Quotes
Greatest_Values


'Routines Finish here

Next i

Sheets(1).Select

End Sub

Sub OrderByDate()

'Double Check and confirm sort by ticker and date

 LastRow = Cells(Rows.Count, 1).End(xlUp).Row
LastCol = Cells(1, Columns.Count).End(xlToLeft).Column

'
    Selection.CurrentRegion.Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range("A2:A" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range("B2:B" & LastCol), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:G" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
End Sub

Sub Get_Quotes()

' Set an initial variable for holding the ticker name, open value and close value
  Dim Ticker_name As String
  Dim Ticker_open As Double
  Dim Ticker_Close As Double

  ' Set an initial variable for holding the volume per Ticker
  Dim Total_vol As Double
  Total_vol = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2


  ' find last row of the data for the loop
  LastRow = Cells(Rows.Count, 2).End(xlUp).Row

  ' Loop through all stock values
  For i = 2 To LastRow


' Check if we are first value per ticker, if it is not...
    If (Cells(i - 1, 1).Value <> Cells(i, 1).Value And Cells(i + 1, 1).Value = Cells(i, 1).Value) Then
    
        'set ticker_open Value
        Ticker_open = Cells(i, 3).Value
        Ticker_name = Cells(i, 1).Value
    End If
    
    
      ' Add to the Total volume per ticker
      Total_vol = Total_vol + Cells(i, 7).Value

' Check if there is the last value per ticker, if it is not...
    If (Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i - 1, 1).Value = Cells(i, 1).Value) Then
    
      'set ticker_close Value
       Ticker_Close = Cells(i, 6).Value
    
    
   On Error Resume Next
   


      ' Print the Ticker data in the Summary Table
      'Metodo de ubicar en una celda
        Range("J1").Value = "Ticker"
        Range("J" & Summary_Table_Row).Value = Ticker_name

        Range("K1").Value = "Yearly Change"
        Range("K" & Summary_Table_Row).Value = Ticker_Close - Ticker_open
        difference = Ticker_Close - Ticker_open
        
        If difference > 0 Then
        
      ' Print the Yearly change to the Summary Table
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Range("K" & Summary_Table_Row).Font.ColorIndex = 1
        
        Else
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        Range("K" & Summary_Table_Row).Font.ColorIndex = 2

        End If
      ' Print the Yearly percent change to the Summary Table
        Range("L1").Value = "Percent"
        Range("L" & Summary_Table_Row).Value = (Ticker_Close / Ticker_open)
        Range("L" & Summary_Table_Row).NumberFormat = "0.00%"


      ' Print the Yearly change to the Summary Table
        Range("M1").Value = "Total Stock Volume"
        Range("M" & Summary_Table_Row).Value = Total_vol
        
        'Resetting to 0 de total volume
        Total_vol = 0


      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
 
End If

Next i


End Sub


Sub clear_data()

'Delete data in target cells for summary
Columns("J:M").Select
Selection.Delete Shift:=xlToLeft

End Sub


Sub Greatest_Values()

Dim LRow As Double


Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"

LRow = Cells(Rows.Count, 10).End(xlUp).Row - 2

Cells(2, 18).Select
ActiveCell.FormulaR1C1 = "=MAX(RC[-6]:R[" & LRow & "]C[-6])"
Cells(2, 17).Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],CHOOSE({2,1},RC[-7]:R[" & LRow & "]C[-7],RC[-5]:R[" & LRow & "]C[-5]),2,0)"


Cells(3, 18).Select
ActiveCell.FormulaR1C1 = "=MIN(R[-1]C[-6]:R[" & LRow - 1 & "]C[-6])"
Cells(3, 17).Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],CHOOSE({2,1},R[-1]C[-7]:R[" & LRow - 1 & "]C[-7],R[-1]C[-5]:R[" & LRow - 1 & "]C[-5]),2,0)"

Cells(4, 18).Select
ActiveCell.FormulaR1C1 = "=MAX(R[-2]C[-5]:R[" & LRow - 2 & "]C[-5])"
Cells(4, 17).Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],CHOOSE({2,1},R[-2]C[-7]:R[" & LRow & "]C[-7],R[-2]C[-4]:R[" & LRow & "]C[-4]),2,0)"

End Sub