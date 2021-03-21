Attribute VB_Name = "Module1"
Sub ticker_metrics_alphabetical_testing()

' define variables: ticker, starting_value, start_date, ending_value, end_date ticker_volume

Dim ticker As String
Dim starting_value As Double
Dim start_date As Integer
Dim ending_value As Double
Dim end_date As Integer
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

' initial conditions

sheets_count = Application.Sheets.Count
ticker_counter = 1
ticker_volume = 0
greatest_percentage_gain_test = 0
greatest_percentage_loss_test = 0
greatest_volume_test = 0

' loop structure

For s = 1 To sheets_count - 1
    
    For t = 2 To Worksheets(s).Cells(Rows.Count, 1).End(xlUp).Row
        
        If t = 2 Then
        starting_value = Worksheets(s).Cells(t, 3).Value
        Worksheets("Collated").Cells(ticker_counter + 1, 2).Value = starting_value
        ticker_volume = CLngLng(Worksheets(s).Cells(t, 7).Value)
        End If
        
        If Worksheets(s).Cells(t + 1, 3) > 0 Then
            If Worksheets(s).Cells(t, 1).Value = Worksheets(s).Cells(t + 1, 1).Value Then
                ticker_volume = ticker_volume + CLngLng(Worksheets(s).Cells(t + 1, 7).Value)
            ElseIf Worksheets(s).Cells(t, 1).Value <> Worksheets(s).Cells(t + 1, 1).Value Then
                ticker_counter = ticker_counter + 1
                ticker = Worksheets(s).Cells(t, 1).Value
                Worksheets("Collated").Cells(ticker_counter, 1).Value = ticker
                ending_value = Worksheets(s).Cells(t, 6).Value
                Worksheets("Collated").Cells(ticker_counter, 3).Value = ending_value
                ticker_delta = ending_value - starting_value
                Worksheets("Collated").Cells(ticker_counter, 4).Value = ticker_delta
                delta_percentage = ticker_delta / starting_value
                Worksheets("Collated").Cells(ticker_counter, 5).Value = delta_percentage
                Worksheets("Collated").Cells(ticker_counter, 6) = ticker_volume
                
                If delta_percentage > greatest_percentage_gain_test Then
                greatest_percentage_gain_test = delta_percentage
                Worksheets("Collated").Cells(2, 9) = ticker
                End If
                Worksheets("Collated").Cells(2, 10).Value = greatest_percentage_gain_test
                
                If delta_percentage < greatest_percentage_loss_test Then
                greatest_percentage_loss_test = delta_percentage
                Worksheets("Collated").Cells(3, 9) = ticker
                End If
                Worksheets("Collated").Cells(3, 10).Value = greatest_percentage_loss_test
                
                If ticker_volume > greatest_volume_test Then
                greatest_volume_test = ticker_volume
                Worksheets("Collated").Cells(4, 9) = ticker
                End If
                Worksheets("Collated").Cells(4, 10).Value = greatest_volume_test
                
                ' conditional formatting
                
                Dim ticker_delta_range: Set ticker_delta_range = Worksheets(s).Cells(ticker_counter, 4)
                Dim gain As FormatCondition
                Dim loss As FormatCondition
                Set gain = ticker_delta_range.FormatConditions.Add(xlCellValue, xlGreater, 0)
                Set loss = ticker_delta_range.FormatConditions.Add(xlCellValue, xlLess, 0)
                
                With gain
                .Interior.Color = vbGreen
                End With
                With loss
                .Interior.Color = vbRed
                End With
                
                starting_value = Worksheets(s).Cells(t + 1, 3)
                Worksheets("Collated").Cells(ticker_counter + 1, 2).Value = starting_value
                ticker_volume = CLngLng(Worksheets(s).Cells(t + 1, 7).Value)
            End If
        End If
        
        ticker = ""
        
    Next t
    
Next s

' conditional formatting
'Dim ticker_delta_range: Set ticker_delta_range = Worksheets("Collated").Cells(4).EntireColumn
'Dim gain As FormatCondition
'Dim loss As FormatCondition
'Set gain = ticker_delta_range.FormatConditions.Add(xlCellValue, xlGreater, 0)
'Set loss = ticker_delta_range.FormatConditions.Add(xlCellValue, xlLess, 0)
'
'With gain
'.Interior.Color = vbGreen
'End With
'With loss
'.Interior.Color = vbRed
'End With

'bonus
'Dim ticker_with_greatest_percentage_increase As String
'Dim greatest_percentage_gain As Double
'Dim greatest_percentage_loss As Double
'Dim greatest_volume As LongLong
'Dim percentage_range: Set percentage_range = Worksheets("Collated").Cells(5).EntireColumn
'Dim volume_range: Set volume_range = Worksheets("Collated").Cells(6).EntireColumn
'greatest_percentage_gain = Application.WorksheetFunction.Max(percentage_range)
'greatest_percentage_loss = Application.WorksheetFunction.Min(percentage_range)
'greatest_volume = Application.WorksheetFunction.Max(volume_range)
'Worksheets("Collated").Cells(2, 9).Value = greatest_percentage_gain
'Worksheets("Collated").Cells(3, 9).Value = greatest_percentage_loss
'Worksheets("Collated").Cells(4, 9).Value = greatest_volume

End Sub

