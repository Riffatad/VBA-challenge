Attribute VB_Name = "Module1"
Sub Quater()

'Declare variables'


Dim tick As String
Dim Volume As Double

Dim openyear As Double
Dim Quatchange As Double
Dim Perchange As Double
Dim sum_table As Integer

Dim Gt_perent_inc As Double
Dim Gt_percent_dec As Double
Dim Gt_t_volume As Double

Dim sumtab As Integer
Dim tick1 As String
Dim tick2 As String
Dim tick3 As String
Dim Quat As Worksheet

'Loop to go through all the sheets in workbook'

For Each Quat In ThisWorkbook.Worksheets

    sum_table = 2
    
    openyear = Quat.Cells(2, 3).Value
   '2nd loop to find the values and moving it into the varibles'
   
   
    For i = 2 To Quat.UsedRange.Rows.Count
    
    

        If Quat.Cells(i + 1, 1).Value <> Quat.Cells(i, 1).Value Then
    
            tick = Quat.Cells(i, 1).Value
            
            closeyear = Quat.Cells(i, 6).Value
            Quatchange = closeyear - openyear
            Perchange = Quatchange / openyear
            Volume = Volume + Quat.Cells(i, 7).Value
   
        
        
        
                Quat.Cells(1, 10).Value = "ticker"
        
                Quat.Cells(sum_table, 10).Value = tick
        
                Quat.Cells(1, 11).Value = "Quaterly Change"
        
                Quat.Cells(sum_table, 11).Value = Quatchange
    
                Quat.Cells(1, 12).Value = "Percent Change"
        
                Quat.Cells(sum_table, 12).Value = Perchange
                Quat.Cells(sum_table, 12).NumberFormat = "0.00%"
                'Columns("L").NumberFormat = "0.00%"'
                
                Quat.Cells(1, 13).Value = "Total Stock Volume"
        
                Quat.Cells(sum_table, 13).Value = Volume
                
          
          'format columns colors based on values '
          
          
            If Quatchange > 0 Then
                Quat.Cells(sum_table, 11).Interior.Color = vbGreen
            Else
                Quat.Cells(sum_table, 11).Interior.Color = vbRed '
            End If
            
            
            sum_table = sum_table + 1
            Volume = 0
            openyear = Quat.Cells(i + 1, 3).Value

        Else
    
            Volume = Volume + Quat.Cells(i, 7).Value
        
        End If
    Next i



    Gt_perent_inc = 0
    Gt_percent_dec = 0
    Gt_t_volume = 0
    
    sumtab = 2

    
              
           
        For i = 2 To sum_table
   
                If Quat.Cells(i, "L").Value > Gt_perent_inc Then
        
                        Gt_perent_inc = Quat.Cells(i, "L").Value
                    
                        tick1 = Quat.Cells(i, "J").Value
                         
                End If
        
                If Quat.Cells(i, "L").Value < Gt_percent_dec Then
        
                     Gt_percent_dec = Quat.Cells(i, "L").Value
                     
                      tick2 = Quat.Cells(i, "J").Value
                    
                     
                End If
            
            
                If Quat.Cells(i, "M").Value > Gt_t_volume Then
        
                Gt_t_volume = Quat.Cells(i, "M").Value
                     
                      tick3 = Quat.Cells(i, "J").Value
                    
                     
                End If
            
        
        Next i
            
            Quat.Cells(1, "Q").Value = "Ticker"
            Quat.Cells(3, "Q").Value = tick1
            Quat.Cells(5, "Q").Value = tick2
            Quat.Cells(7, "Q").Value = tick3
        
            Quat.Cells(1, "R").Value = "Value"
            
            Quat.Cells(3, "O").Value = "Greatest % increase"
            Quat.Cells(3, "R").Value = Gt_perent_inc
            Quat.Cells(5, "O").Value = "Greatest % decrease"
            Quat.Cells(5, "R").Value = Gt_percent_dec
            Quat.Cells(7, "O").Value = "Greatest total volume"
            Quat.Cells(7, "R").Value = Gt_t_volume
        
        
        
         
     
    



        
    Next Quat

End Sub

