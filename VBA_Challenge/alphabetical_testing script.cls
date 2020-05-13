VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabetical_testing_V2()
    
        ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
    
        'Set variables for holding Ticker, Yearly Change, Percent Change, and Total Stock Volume
        Dim Ticker As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Dim TickerNum As Long
        Dim Start_Value As Double
        Dim End_Value As Double

        Yearly_Change = 0
        Percent_Change = 0
        Total_Stock_Volume = 0
        Start_Value = Cells(2, 3).Value
        End_Value = 0
        
        'Keep track of the location for each Ticker
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
            
        'Determine the last Row
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Loop through rows
        For i = 2 To LastRow
    
        ' Check if we are still within the same ticker, if it is not then identify next ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the Ticker
            Ticker = Cells(i, 1).Value
            
            'Determine end value
            End_Value = Cells(i, 6).Value
            
            'Calc the Yearly Change
            Yearly_Change = End_Value - Start_Value
            
            'Calc Percent Change
            If Start_Value <> 0 Then
                Percent_Change = Yearly_Change / Start_Value
            Else
                Percent_Change = 0
            End If
            
            'Calc Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            'Add one to the ticker
            TickerNum = TickerNum + 1
            
            'Print the Ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker
    
            'Print the Yearly Change to the Summary Table
            Range("J" & Summary_Table_Row).Value = Yearly_Change
                
            'Print the Ticker in the Summary Table
            Range("K" & Summary_Table_Row).Value = Percent_Change
    
            'Print the Yearly Change to the Summary Table
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset the Start Value
            Start_Value = Cells(i + 1, 3).Value
            
            'Reset the End Value
            End_Value = 0
            
            'Reset the Yearly Change
            Yearly_Change = 0
            
            'Reset the Percent Change
            Percent_Change = 0
            
            'Reset the Stock_Volume
            Total_Stock_Volume = 0
            
        'If the cell immediately following a row is the same ticker
        Else
            
            'Calc the Yearly Change
            Yearly_Change = Cells(i, 6).Value - Cells(i, 3).Value
            
            'Calc Percent Change
            If Start_Value <> 0 Then
                Percent_Change = Yearly_Change / Start_Value
            Else
                Percent_Change = 0
            End If
            
            'Calc Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
       End If
       
       Next i
       
       'Format Yearly Change Color
       'Determine the last Row
        LastColorRow = Cells(Rows.Count, "J").End(xlUp).Row
        
        For i = 2 To LastColorRow
                 
            If (Cells(i, 10).Value >= 0) Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
            
         Next i
                
        'Display Totals
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
             
    
    'Move to next sheet
    Next ws
    
End Sub



