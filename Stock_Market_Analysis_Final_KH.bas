Attribute VB_Name = "Stock_Market_Analysis_Final_KH"

' Stock Market Analyst

' 1. Add row headers for Ticker, Yearly Change, Percent Chage, Total Stock Volume and Challenges
' 2. Add Ticker symbol
' 3. Calculate the yearly change from opening price at the beginning of the year to the closing price at the end of the year
' 4. Calculate the percent change from the opening price at the beginning of a given year to the closing price at the end of the year
' 5. Calculate the total stock volume of the stock
' 6. Add conditional formatting that will highlight the positive change in green and negative change in red
' 7. Add Challenges to show maximum & minimum percent change and total volume increase
' 8. Run subroutine for all sheets


Sub Stock_Market_Analysis()

Dim xSh As Worksheet
For Each xSh In Worksheets
    xSh.Select
    

    ' Add header titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Set an initial variable for outputs
    Dim Ticker As String
    Dim Open_Price As Double
    Open_Price = 0
    Dim Closing_Price As Double
    Closing_Price = 0
    Dim Change_Price As Double
    Change_Price = 0
    Dim Percent_Change As Double
    Percent_Change = 0
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim Max_Ticker_Name As String
    Max_Ticker_Name = " "
    Dim Min_Ticker_Name As String
    Min_Ticker_Name = " "
    Dim Max_Percent As Double
    Max_Percent = 0
    Dim Min_Percent As Double
    Min_Percent = 0
    Dim Max_Volume_Ticker As String
    Max_Volume_Ticker = " "
    Dim Max_Volume As Double
    Max_Volume = 0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' Keep track of the location of each Ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Determine the last row
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Determine Open_Price of first Ticker
    Open_Price = Cells(2, 3).Value
    
    ' Add output information
    For i = 2 To LastRow
    
        ' Check if we are still within the same Ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ' Set Ticker name
            Ticker = Cells(i, 1).Value
            
            ' Add to the Total_Stock_Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
                ' Calculate the Closing_Price, Change_Price, and Percent_Change
                Closing_Price = Cells(i, 6).Value
                Change_Price = Closing_Price - Open_Price
                'Check Division by 0 condition
                If Open_Price <> 0 Then
                Percent_Change = (Change_Price / Open_Price) * 100
                End If
            
            ' Print the Ticker name in the output
            Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Print Change_Price to output table
            Range("J" & Summary_Table_Row).Value = Change_Price
                If (Change_Price > 0) Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Change_Price < 0) Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
            ' Print the Percent_Change
            Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
            
            ' Print the Total_Stock_Volume to the output
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the Prices
            Change_Price = 0
            Closing_Price = 0
            Open_Price = Cells(i + 1, 3).Value
            
            ' Challenges
            If (Percent_Change > Max_Percent) Then
                Max_Percent = Percent_Change
                Max_Ticker_Name = Ticker
            ElseIf (Percent_Change < Min_Percent) Then
                Min_Percent = Percent_Change
                Min_Ticker_Name = Ticker
            End If
            
            If (Total_Stock_Volume > Max_Volume) Then
                Max_Volume = Total_Stock_Volume
                Max_Volume_Ticker = Ticker
            End If
            
            ' Print Challenges
            Cells(2, 16).Value = Max_Ticker_Name
            Cells(3, 16).Value = Min_Ticker_Name
            Cells(4, 16).Value = Max_Volume_Ticker
            Cells(2, 17).Value = (CStr(Max_Percent) & "%")
            Cells(3, 17).Value = (CStr(Min_Percent) & "%")
            Cells(4, 17).Value = Max_Volume
            Columns("O:O").EntireColumn.AutoFit
            
            ' Reset Percent_Change
            Percent_Change = 0
            
            ' Reset the Total_Stock_Volume
            Total_Stock_Volume = 0
            
            ' If the cell immediately following a row is the same Ticker
            Else
            
                ' Add to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
        End If
        
    Next i
    
Next xSh

    
End Sub

