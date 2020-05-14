VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()
   
    'Define varibles
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock As LongLong
    Dim First_Date As Long
    Dim Last_Date As Long
    Dim First_Open As Double
    Dim Last_Close As Double
    Dim Summary_Table_Row As Long
    Dim LastRow As Long
    Dim WS As Worksheet

    'Loop through each worksheet
    For Each WS In Worksheets
    
        'For each worksheet, reset counter values
        Total_Stock = 0
        Summary_Table_Row = 2
        
        'For each worksheet find first and last dates for the year
        First_Date = Application.WorksheetFunction.Min(WS.Columns("B"))
        Last_Date = Application.WorksheetFunction.Max(WS.Columns("B"))

        'For each worksheet, populate headers for summary table and set column width
        WS.Range("I1").Value = "Ticker"
        WS.Columns("I").ColumnWidth = 10
        WS.Range("J1").Value = "Yearly Change"
        WS.Columns("J").ColumnWidth = 12
        WS.Range("K1").Value = "Percent Change"
        WS.Columns("K").ColumnWidth = 13
        WS.Range("L1").Value = "Total Stock Volume"
        WS.Columns("L").ColumnWidth = 16
    
        'For each worksheet, find the last row of data
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through all rows in the worsheet
        For i = 2 To LastRow
    
            'Check to see if Tickers match and if they don't, then do the following steps
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

                'Populate the ticker in the summary table
                Ticker = WS.Cells(i, 1).Value
                WS.Range("I" & Summary_Table_Row).Value = Ticker

                'If this is the first date of the year, use this first open date amount
                If WS.Cells(i, 2).Value = First_Date Then
                    First_Open = WS.Cells(i, 3).Value
                    End If

                'If this is the last date of the year, use this last close date amount
                If WS.Cells(i, 2).Value = Last_Date Then
                    Last_Close = WS.Cells(i, 6).Value
                    End If

                'Calculate the yearly change amount and populate summary table
                Yearly_Change = Last_Close - First_Open
                WS.Range("J" & Summary_Table_Row).Value = Yearly_Change

                'Conditional formatting for yearly change amount
                If Yearly_Change > 0 Then
                    WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf Yearly_Change < 0 Then
                    WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
                'Calculate the percent change for the year; if it is zero, set to NA since can't divide by zero; populate summary table
                If First_Open = 0 Then
                    WS.Range("K" & Summary_Table_Row).Value = "NA"
                ElseIf First_Open > 0 Then
                    Percent_Change = (Last_Close - First_Open) / First_Open
                    WS.Range("K" & Summary_Table_Row).Value = Percent_Change
                    WS.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                End If

                'Add the last total stock volume to amount and populate summary table
                Total_Stock = Total_Stock + WS.Cells(i, 7).Value
                WS.Range("L" & Summary_Table_Row).Value = Total_Stock
 
                'Increment counter for rows
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Set Total Stock Volume counter back to zero
                Total_Stock = 0
        
            Else
        
                'If Tickers match, then do the following steps
                'If this is the first date of the year, use this first open date amount
                If WS.Cells(i, 2).Value = First_Date Then
                    First_Open = WS.Cells(i, 3).Value
                End If

                'If this is the last date of the year, use this last close date amount
                If WS.Cells(i, 2).Value = Last_Date Then
                    Last_Close = WS.Cells(i, 6).Value
                End If

                'Add to the last total stock volume sum
                Total_Stock = Total_Stock + WS.Cells(i, 7).Value

            End If
        
        Next i
        
    Next WS
          
    'Challenge Code

    'Define Variables
    Dim Total_Vol As LongLong
    Dim Percent_Inc As Double
    Dim Percent_Dec As Double
    Dim NewLastRow As Integer

    'Loop through each worksheet
    For Each WS In Worksheets
        
        'For each worksheet, find last row in summary table created above
        NewLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row

        'For each worksheet, determine the values for Greatest Percent Increase, Greatest Percent Decrease and Greatest Stock Volume
        Percent_Inc = Application.WorksheetFunction.Max(WS.Columns("K"))
        Percent_Dec = Application.WorksheetFunction.Min(WS.Columns("K"))
        Total_Vol = Application.WorksheetFunction.Max(WS.Columns("L"))
  
        'For each worksheet, populate headers and row lables for new summary table and set column widths
        WS.Range("O1").Value = "Ticker"
        WS.Columns("O").ColumnWidth = 10
        WS.Range("P1").Value = "Value"
        WS.Columns("P").ColumnWidth = 16
        WS.Range("N2").Value = "Greatest % Increase"
        WS.Range("N3").Value = "Greatest % Decrease"
        WS.Range("N4").Value = "Greatest Total Volume"
        WS.Columns("N").ColumnWidth = 19

        'Loop through each row of prior summary table
        For i = 2 To NewLastRow

            'Find ticker with Greatest Percent Increase and put in new summary table
            If WS.Cells(i, 11) = Percent_Inc Then
                WS.Range("O2").Value = WS.Cells(i, 9).Value
                WS.Range("P2").Value = WS.Cells(i, 11).Value
                WS.Range("P2").NumberFormat = "0.00%"

            'Find ticker with Greatest Percent Decrease and put in new summary table
            ElseIf WS.Cells(i, 11) = Percent_Dec Then
                WS.Range("O3").Value = WS.Cells(i, 9).Value
                WS.Range("P3").Value = WS.Cells(i, 11).Value
                WS.Range("P3").NumberFormat = "0.00%"
            End If

            'Find ticker with Greatest Stock Volume and put in new summary table
            If WS.Cells(i, 12) = Total_Vol Then
                WS.Range("O4").Value = WS.Cells(i, 9).Value
                WS.Range("P4").Value = WS.Cells(i, 12).Value
            End If

        Next i
        
    Next WS

End Sub


