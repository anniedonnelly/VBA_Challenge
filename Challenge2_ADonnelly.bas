Attribute VB_Name = "Module1"
Sub MultiYearStockData()
   
For Each ws In Worksheets
 
    'Dimension Variables
    Dim WSName As String
        WSName = ws.Name
    Dim Ticker As String
    Dim YearChange As Double
        YearChange = 0
    Dim YearOpen As Double
        YearOpen = ws.Range("C2").Value
    Dim YearClose As Double
        YearClose = 0
    Dim SummaryRow As Integer
        SummaryRow = 2
    Dim StockVol As Double
        StockVol = 0
        PerChange = 0
    Dim PerChangeOpen As Double
        PerChangeOpen = ws.Range("C2").Value
    Dim PerChangeClose As Double
        PerChangeClose = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Table Titles
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Yearly Change"
    ws.Range("L1") = "Percent Change"
    ws.Range("M1") = "Total Stock Volume"
    
    'Bonus Table Diminsion Variables
    Dim GreatIn As Double
        GreatIn = 0
    Dim InTicker As String
    Dim GreatDe As Double
        GreatDe = 0
    Dim DeTicker As String
    Dim GreatTotal As Double
        GreatTotal = 0
    Dim TotTicker As String
     
    
    'Loop = ticker and yearly total
    For i = 2 To LastRow
        
        'If asking for same ticker, if not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Set Ticker name
            Ticker = ws.Cells(i, 1).Value
            ws.Range("J" & SummaryRow).Value = Ticker
            
            'Set YearChange
            YearClose = ws.Cells(i, 6).Value
            YearChange = YearClose - YearOpen
            ws.Range("K" & SummaryRow) = YearChange
                'conditional formating
                If YearChange > 0 Then
                    ws.Range("K" & SummaryRow).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & SummaryRow).Interior.ColorIndex = 3
                End If
            YearChange = 0
            YearOpen = ws.Cells(i + 1, 3)
            
            'Set Percent Change
            PerChangeClose = ws.Cells(i, 6).Value
            PerChange = PerChangeClose - PerChangeOpen
            PerChange = PerChange / PerChangeOpen
                'Finding the Greatest increase
                If PerChange > GreatIn Then
                    GreatIn = PerChange
                    InTicker = ws.Cells(i, 1).Value
                Else
                    GreatIn = GreatIn
                    InTicker = InTicker
                End If
                'Finding the greatest decrease
                If GreatDe > PerChange Then
                    GreatDe = PerChange
                    DeTicker = ws.Cells(i, 1).Value
                Else
                    GreatDe = GreatDe
                    DeTicker = DeTicker
                End If
            PerChange = Format(PerChange, "0.0000%")
            ws.Range("L" & SummaryRow) = PerChange
            PerChange = 0
            PerChangeOpen = ws.Cells(i + 1, 3)
            
            'Set Stock Volume
            StockVol = StockVol + ws.Cells(i, 7).Value
                'Finding the Greatest Stock Volume
                If StockVol > GreatTotal Then
                    GreatTotal = StockVol
                    TotTicker = ws.Cells(i, 1).Value
                Else
                    GreatTotal = GreatTotal
                    TotTicker = TotTicker
                End If
            ws.Range("M" & SummaryRow) = StockVol
            StockVol = 0
            
            'Adding to Table row
            SummaryRow = SummaryRow + 1
        
        'If Is the same ticker
        Else
            
            'Adding to Stock Volume
            StockVol = StockVol + Cells(i, 7).Value
        
        End If
        
    Next i
    
    'Bonus Table
    ws.Range("P2") = "Greatest % Increase"
    ws.Range("P3") = "Greatest % Decrease"
    ws.Range("P4") = "Greatest Total Volume"
    ws.Range("Q1") = "Ticker"
    ws.Range("R1") = "Value"
    
    'Inputing into bonus table
    ws.Range("Q2").Value = InTicker
    ws.Range("Q3").Value = DeTicker
    ws.Range("Q4").Value = TotTicker
    ws.Range("R2").Value = Format(GreatIn, "0.00%")
    ws.Range("R3").Value = Format(GreatDe, "0.00%")
    ws.Range("R4").Value = GreatTotal
    
Next ws

End Sub

