Attribute VB_Name = "Module1"
Sub stock()
 
 For Each ws In Worksheets
       
        TotalVol = 0
        cPointer = 2
        iPointer = 2
        
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To RowCount
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
            
              TotalVol = TotalVol + ws.Cells(i, "G").Value
              
                StartPrice = ws.Cells(cPointer, "C").Value
                EndPrice = ws.Cells(i, "F").Value
                YearlyChange = StartPrice - EndPrice
                
                ws.Cells(iPointer, "I").Value = ws.Cells(i, "A").Value
                ws.Cells(iPointer, "J").Value = YearlyChange
                ws.Cells(iPointer, "K").Value = "%" & (YearlyChange / StartPrice * 100)
                ws.Cells(iPointer, "L").Value = TotalVol
                
                 TotalVol = 0
                cPointer = i + 1
                iPointer = iPointer + 1
                
           Else
                TotalVol = TotalVol + ws.Cells(i, "G").Value
                
                End If

        Next i

    Next ws
   
    MsgBox ("Done")

End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Save
End Sub
