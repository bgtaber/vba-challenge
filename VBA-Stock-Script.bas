Attribute VB_Name = "Module1"
Sub year_loop():
    'declare variables
    Dim ticker As String
    Dim vol As LongLong
    Dim price_open As Double
    Dim price_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim i As Long
    Dim i_sum As Long
    Dim stock_vol As Long
    Dim last_row As Long
    Dim ws_count As Integer
    Dim w As Integer
                
        Range("I1").Value = "sum ticker"
        Range("J1").Value = "price change"
        Range("K1").Value = "percent change"
        Range("L1").Value = "total volume"
        
 
         ws_count = ActiveWorkbook.Worksheets.Count
     
         
         For w = 1 To ws_count
         
         last_row = Cells(Rows.Count, 1).End(xlUp).Row
         i_sum = 2
         yearly_change = 0
         Start = 2
            'iterate data
            For i = 2 To last_row
                ticker = Cells(i, 1).Value
                vol = vol + Cells(i, 7).Value
            
                If Cells(Start, 3).Value = 0 Then
                    For find_value = Start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            Start = find_value
                            Exit For
                    End If
                        Next find_value
                End If
            
                
                
                If Cells(i + 1, 1) <> ticker Then
                    yearly_change = (Cells(i, 6) - Cells(Start, 3))
                    percent_change = Round(((yearly_change) / Cells(Start, 3)) * 100, 2)
                    'output ticker
                    Cells(i_sum, 9).Value = ticker
                    'calc total volume for each ticker group
                    Cells(i_sum, 12).Value = vol
                    Cells(i_sum, 10).Value = yearly_change
                        If yearly_change < 0 Then
                            Cells(i_sum, 10).Interior.Color = vbRed
                        Else
                            Cells(i_sum, 10).Interior.Color = vbGreen
                        End If
                    'calc percent_change
                    Cells(i_sum, 11).Value = percent_change
                    'Cells(i_sum, 11).NumberFormat = "0.00%"

                    i_sum = i_sum + 1
                    yearly_change = 0
                    vol = 0
                    Start = i + 1
                    
                          
                End If
            Next i
            'MsgBox ActiveWorkbook.Worksheets(w).Name
        Next w
        
End Sub

