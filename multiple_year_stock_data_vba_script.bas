Attribute VB_Name = "Module1"
Sub ticker_metrics_multiple_year_stock_data()

' define variables

Dim ticker_header As String
Dim ticker As String
Dim starting_value As Double
Dim ending_value As Double
Dim ticker_delta As Double
Dim delta_percentage As Double
Dim ticker_volume As LongLong
'bonus
Dim greatest_percentage_gain_test As Double
Dim greatest_percentage_loss_test As Double
Dim greatest_volume_test As LongLong

' loop automating variables

Dim sheets_count As Integer
Dim ticker_counter As Integer
Dim starting_value_error As Integer
Dim gain As FormatCondition
Dim loss As FormatCondition

' initial conditions

Cells.FormatConditions.Delete
sheets_count = Application.Sheets.Count
ticker_counter = 1
ticker_volume = 0
greatest_percentage_gain_test = 0
greatest_percentage_loss_test = 0
greatest_volume_test = 0
ticker_header = "Ticker"

' loop structure

For s = 1 To sheets_count

    Worksheets(s).Cells(1, 9).Value = ticker_header
    Worksheets(s).Cells(1, 15).Value = ticker_header
    Worksheets(s).Cells(1, 10).Value = "Yearly Change"
    Worksheets(s).Cells(1, 11).Value = "Percent Change"
    Worksheets(s).Cells(1, 12).Value = "Total Stock Volume"
    Worksheets(s).Cells(1, 16).Value = "Value"
    Worksheets(s).Cells(2, 14).Value = "Greatest % Increase"
    Worksheets(s).Cells(3, 14).Value = "Greatest % Decrease"
    Worksheets(s).Cells(4, 14).Value = "Greatest Total Volume"
    
    For t = 2 To Worksheets(s).Cells(Rows.Count, 1).End(xlUp).Row
        
        If t = 2 Then
        starting_value = CDbl(Worksheets(s).Cells(t, 3).Value)
        ticker_volume = CLngLng(Worksheets(s).Cells(t, 7).Value)
        End If
        
        If Worksheets(s).Cells(t + 1, 3) > 0 Then
            If Worksheets(s).Cells(t, 1).Value = Worksheets(s).Cells(t + 1, 1).Value Then
                ticker_volume = ticker_volume + CLngLng(Worksheets(s).Cells(t + 1, 7).Value)
            ElseIf Worksheets(s).Cells(t, 1).Value <> Worksheets(s).Cells(t + 1, 1).Value Then
                ticker_counter = ticker_counter + 1
                ticker = Worksheets(s).Cells(t, 1).Value
                Worksheets(s).Cells(ticker_counter, 9).Value = ticker
                'Worksheets(s).Cells(ticker_counter, 13).Value = starting_value
                ending_value = CDbl(Worksheets(s).Cells(t, 6).Value)
                'Worksheets(s).Cells(ticker_counter, 14).Value = ending_value
                ticker_delta = CDbl(ending_value - starting_value)
                Worksheets(s).Cells(ticker_counter, 10).Value = ticker_delta
                
                If starting_value_error = 1 Then
                ticker_delta = 0
                delta_percentage = ticker_delta
                ElseIf starting_value_error = 0 Then
                delta_percentage = CDbl(ticker_delta / starting_value)
                End If
                
                'If starting_value <> 0 Then
                'delta_percentage = CDbl(ticker_delta / starting_value)
                'ElseIf starting_value_error = 0 Then
                'delta_percentage = 0
                'End If
                
                Worksheets(s).Cells(ticker_counter, 11).Value = delta_percentage
                Worksheets(s).Cells(ticker_counter, 12) = ticker_volume
                
                If delta_percentage > greatest_percentage_gain_test Then
                greatest_percentage_gain_test = delta_percentage
                Worksheets(s).Cells(2, 15) = ticker
                End If
               
                If delta_percentage < greatest_percentage_loss_test Then
                greatest_percentage_loss_test = delta_percentage
                Worksheets(s).Cells(3, 15) = ticker
                End If
                
                If ticker_volume > greatest_volume_test Then
                greatest_volume_test = ticker_volume
                Worksheets(s).Cells(4, 15) = ticker
                End If
                
                ticker_volume = CLngLng(Worksheets(s).Cells(t + 1, 7).Value)
                
                If Worksheets(s).Cells(t + 1, 3) > 0 Then
                starting_value = CDbl(Worksheets(s).Cells(t + 1, 3))
                starting_value_error = 0
                ElseIf Worksheets(s).Cells(t + 1, 3) = 0 Then
                starting_value = 0
                starting_value_error = 1
                End If
            End If
        End If
        
        ticker = ""
        
    Next t
    
' conditional formatting
Dim ticker_delta_range: Set ticker_delta_range = Range("j2:j" & Worksheets(s).Cells(Rows.Count, 1).End(xlUp).Row)
Set gain = ticker_delta_range.FormatConditions.Add(xlCellValue, xlGreater, 0)
Set loss = ticker_delta_range.FormatConditions.Add(xlCellValue, xlLess, 0)

With gain
.Interior.Color = vbGreen
End With
With loss
.Interior.Color = vbRed
End With

Dim delta_percentage_range: Set delta_percentage_range = Worksheets(s).Cells(11).EntireColumn
With delta_percentage_range
.NumberFormat = "0.00%"
End With

Worksheets(s).Cells(2, 16).Value = greatest_percentage_gain_test
Dim greatest_percentage_gain_test_range: Set greatest_percentage_gain_test_range = Worksheets(s).Cells(2, 16)
With greatest_percentage_gain_test_range
.NumberFormat = "0.00%"
End With

Worksheets(s).Cells(3, 16).Value = greatest_percentage_loss_test
Dim greatest_percentage_loss_test_range: Set greatest_percentage_loss_test_range = Worksheets(s).Cells(3, 16)
With greatest_percentage_loss_test_range
.NumberFormat = "0.00%"
End With

Worksheets(s).Cells(4, 16).Value = greatest_volume_test
Dim greatest_volume_test_range: Set greatest_volume_test_range = Worksheets(s).Cells(4, 16)
With greatest_volume_test_range
.NumberFormat = "General"
End With

'reset s-loop initial conditions
ticker_counter = 1
greatest_percentage_gain_test = 0
greatest_percentage_loss_test = 0
greatest_volume_test = 0

Next s

End Sub

