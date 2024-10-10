Attribute VB_Name = "Module1"
Option Explicit
Sub Q1toQ4()



    Dim row As Double
    Dim column As Integer
    Dim ws As Worksheet
    Dim ticker As String
    Dim volume As Double
    Dim tickercount As Integer
    Dim opened As Double
    Dim closed As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim greatvolume As Double
    Dim greatincrease As Double
    Dim greatdecrease As Double
    
    
    
    
    
    For Each ws In Worksheets
    
       
            Dim WorksheetName As String
            Dim LastRow As Double
            Dim rng As Range
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
            WorksheetName = ws.Name
            tickercount = 0
            ws.Range("I1,Q1").Value = "Ticker"
            ws.Range("R1").Value = "Value"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("P2").Value = "Greastest % Increase"
            ws.Range("P3").Value = "Greastest % Decrease"
            ws.Range("P4").Value = "Greatest Total Volume"
            
            ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
            ws.Range("R2:R3").NumberFormat = "0.00%"
            
            greatvolume = 0
            greatincrease = 0
            greatdecrease = 0
            
            For row = 2 To LastRow
            
                    
            
                If ws.Cells(row, 1).Value <> ticker Then
                   
                    ticker = ws.Cells(row, 1).Value
                    volume = 0
                    tickercount = tickercount + 1
                    
                    opened = ws.Cells(row, 3).Value
                    
                    
                    
            
                End If
                    
                    
                    volume = volume + ws.Cells(row, 7).Value
            
                If ws.Cells(row + 1, 1).Value <> ticker Then
                    
                    ws.Cells(tickercount + 1, 9).Value = ticker
                    
                    closed = ws.Cells(row, 6)
                    QuarterlyChange = closed - opened
                    
                                                       
                    ws.Cells(tickercount + 1, 10).Value = QuarterlyChange
                    
                    PercentChange = (closed - opened) / opened
                        
                                               
                        
                    ws.Cells(tickercount + 1, 11).Value = PercentChange
                    
                    ws.Cells(tickercount + 1, 12).Value = volume
                        
                                           
                        
                        If QuarterlyChange > 0 Then
                            ws.Cells(tickercount + 1, 10).Interior.ColorIndex = 4
                        
                        
                        ElseIf QuarterlyChange < 0 Then
                            ws.Cells(tickercount + 1, 10).Interior.ColorIndex = 3
                        
                        ElseIf QuarterlyChange = 0 Then
                            ws.Cells(tickercount + 1, 10).Interior.ColorIndex = 0
                        End If
                        
                        
                        If volume > greatvolume Then
                            greatvolume = volume
                            ws.Range("Q4").Value = ticker
                        End If
                        
                        If PercentChange > greatincrease Then
                            greatincrease = PercentChange
                            ws.Range("Q2").Value = ticker
                        
                        ElseIf PercentChange < greatdecrease Then
                            greatdecrease = PercentChange
                            ws.Range("Q3").Value = ticker
                        End If
            
                End If
            
                
                
            
            
            Next row
            
                
                   ws.Range("R2").Value = greatincrease
                   ws.Range("R3").Value = greatdecrease
                   ws.Range("R4").Value = greatvolume
        
    Next ws
    
End Sub


























