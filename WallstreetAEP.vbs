Sub wallstreet()

Dim lastWS As Integer

lastWS = ThisWorkbook.Sheets.Count

For ws = 1 To lastWS

Worksheets(ws).Activate;'

'Add Summary Table Column headings
    Range("I1:N1").Value = Array("Ticker2", "Yearly Change", "Percent Change", "Total Stock Vloume", "Open", "Close")

    'set variables
    Dim Ticker As String
    Dim YearlyChangeJ As Double
    Dim PercentChangeK As Double
    Dim TotalStock As Double
        TotalStock = 0
    Dim OpenM As Double
    Dim CloseN As Double
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
    
    OpenM = Cells(2, 3).Value
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
    
    Ticker = Cells(i, 1).Value
    CloseN = Cells(i, 6).Value
    
    
'is this the first occurence of ticker symbol?
    If Cells(i - 1, 1).Value <> Cells(i, 1) Then
    
        Ticker = Cells(i, 1).Value
        CloseN = Cells(i, 6).Value
    
        Range("M" & SummaryTableRow).Value = OpenM
  
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
        Range("I" & SummaryTableRow).Value = Ticker
        Range("N" & SummaryTableRow).Value = CloseN
    
        TotalStock = TotalStock + Cells(i, 7).Value
        Range("L" & SummaryTableRow).Value = TotalStock
     
        YearlyChangeJ = Cells(i, 10).Value
        YearlyChangeJ = OpenM - CloseN
        Range("J" & SummaryTableRow).Value = YearlyChangeJ

            If Range("J" & SummaryTableRow).Value < 0 Then
                Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            Else
                Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            End If
            
    If OpenM = 0 Then
    PercentChangeK = 0
    
    Else
        
        
    'PercentChangeK = Cells(i, 11).Value
    PercentChangeK = (CloseN - OpenM) / OpenM
    
    End If
    
    OpenM = Cells(i + 1, 3).Value
    
    Range("K" & SummaryTableRow).Value = PercentChangeK
        
        Cells(i, 11).Style = "Percent"
        Cells(i, 11).NumberFormat = "0.00%"

    SummaryTableRow = SummaryTableRow + 1
    TotalStock = 0
    
    
'if the cell immediately following this row is the same ticker:
    Else
    
    TotalStock = TotalStock + Cells(i, 7).Value
    
    End If
    
Next i

Next ws


  
End Sub
