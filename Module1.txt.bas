Attribute VB_Name = "Module1"
Sub VBA_challange1()
Attribute VBA_challange1.VB_ProcData.VB_Invoke_Func = " \n14"

For Each ws In Worksheets
 
 ' Dim unique value() As String
 ' Determine the Last Row in sheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
 ' Determine the Last column in sheet
    lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).column
        
    Dim ticker As String
    Dim column As Integer
    Dim volume As Double
    Dim Total_stock_volume As Long
    Dim yearly_change As Double
    Dim open_price As Double
    Dim percentagechange As Double
    Dim IsFirst As Integer
    Dim first_price As Double

    volume = 0

    column = 1
    rownumber = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To lastrow

        ticker = ws.Cells(i, 1).Value

        ' Searches for when the value of the next cell is different than that of the current cell
        If ws.Cells(i, column).Value <> ws.Cells(i + 1, column).Value Then
            close_price = ws.Cells(i, 6).Value
            yearly_change = first_price - close_price
            If first_price <> 0 Then
                percentagechange = Round((yearly_change / first_price * 100), 2)
                ws.Cells(rownumber, 9).Value = ticker
                ws.Cells(rownumber, 12).Value = volume
                ws.Cells(rownumber, 10).Value = yearly_change
                ws.Cells(rownumber, 11).Value = "%" & percentagechange
            End If
            If (percentagechange > 0) Then
                ws.Cells(rownumber, 11).Font.Color = vbBlack
                ws.Cells(rownumber, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(rownumber, 11).Font.Color = vbBlack
                    ws.Cells(rownumber, 11).Interior.Color = vbRed
 
            End If
            rownumber = rownumber + 1
            volume = 0
            IsFirst = 0
        Else
            IsFirst = IsFirst + 1
            volume = ws.Cells(i + 1, 7).Value + volume
            If IsFirst = 1 Then
            first_price = ws.Cells(i, 3).Value
            End If
    End If
     
    Next i

Next ws

End Sub
