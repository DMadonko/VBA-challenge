Attribute VB_Name = "Module1"
Sub WorksheetsLoop()
    
    Dim CurrentWs As Worksheet
    Dim command_spreadsheet As Boolean

    
    For Each CurrentWs In Worksheets
    
        Dim ticker As String
        ticker = " "
        
    
        Dim open_value As Double
        open_value = 0
        Dim close_value As Double
        close_value = 0
        Dim yearly_change As Double
        yearly_change = 0
        Dim percentage As Double
        percentage = 0
        
        Dim min_ticker As String
        min_ticker = " "
        Dim max_ticker As String
        max_ticker = " "
        Dim min_percentage As Double
        min_percentage = 0
        Dim max_percentage As Double
        max_percentage = 0
       
        Dim max_total_ticker As String
        max_total_ticker = " "
        Dim max_total As Double
        max_total = 0
        
        Dim stocktotal As Double
        stocktotal = 0
        
        
        Dim summary As Long
        summary = 2
        
        
        Dim endrow As Long
        Dim i As Long
        
        endrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        
        open_value = CurrentWs.Cells(2, 3).Value
        
        For i = 2 To Lastrow
        
      
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
               
                ticker = CurrentWs.Cells(i, 1).Value
                
                
                close_value = CurrentWs.Cells(i, 6).Value
                yearly_change = close_value - open_value
               
                If open_value <> 0 Then
                    yearly_change = (yearly_change / open_value) * 100
                End If
                
            
                stocktotal = stocktotal + CurrentWs.Cells(i, 7).Value
              
                
                
                CurrentWs.Range("I" & summary).Value = ticker
                CurrentWs.Range("J" & summary).Value = yearly_change
                If (yearly_change > 0) Then
                    CurrentWs.Range("J" & summary).Interior.ColorIndex = 4
                ElseIf (yearly_change <= 0) Then
                    CurrentWs.Range("J" & summary).Interior.ColorIndex = 3
                End If
                
                
                CurrentWs.Range("K" & summary).Value = (CStr(yearly_change) & "%")
                CurrentWs.Range("L" & summary).Value = stocktotal
                
                
                summary = summary + 1
                
                yearly_change = 0
                
                close_value = 0
              
                open_value = CurrentWs.Cells(i + 1, 3).Value
              
                

                If (yearly_change > max_percentage) Then
                    max_percentage = yearly_change
                    max_ticker = ticker
                ElseIf (yearly_change < min_percentage) Then
                    min_percentage = yearly_change
                    min_ticker = ticker
                End If
                       
                If (stocktotal > max_total) Then
                    max_total = stocktotal
                    max_total_ticker = ticker
                End If
                
                
                yearly_change = 0
                stocktotal = 0
                
            
            Else
                
                stocktotal = stocktotal + CurrentWs.Cells(i, 7).Value
            End If
            
      
        Next i

            If Not command_spreadsheet Then
            
                CurrentWs.Range("Q2").Value = (CStr(max_percentage) & "%")
                CurrentWs.Range("Q3").Value = (CStr(min_percentage) & "%")
                CurrentWs.Range("P2").Value = max_ticker
                CurrentWs.Range("P3").Value = min_ticker
                CurrentWs.Range("Q4").Value = max_total
                CurrentWs.Range("P4").Value = max_total_ticker
                
            Else
                command_spreadsheet = False
            End If
        
     Next CurrentWs
End Sub



