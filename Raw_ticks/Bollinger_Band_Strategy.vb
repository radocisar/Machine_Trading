Public Class Bollinger_Band_Strategy
    Private WithEvents Candles_creation_data_processing_event As New Candles_Creation_Data_Processing

    Public upper_band As Double
    Public lower_band As Double
    Public middle_band As Double
    Public Upper_Lower_Band_Span As Double

    Sub Bollinger_Band_Trading_Strategy(BB_candle_arr(,) As Double, open_price As Double, high_price As Double, low_price As Double, last_price As Double) Handles Candles_creation_data_processing_event.start_trading_strategy

        Dim Prev_Candle_High_and_Close_above_U_Band As Boolean
        Dim Prev_Candle_Low_and_Close_below_L_Band As Boolean
        Dim Prev_low As Double
        Dim Prev_high As Double

#Region "Variables assignment and prep work"

        ' Shift all values by one so that last candle can be updated with latest data (open, high, low, close)
        For n = 0 To Form1.candle_init_count - 2
            BB_candle_arr(n, 0) = BB_candle_arr(n + 1, 0)
            BB_candle_arr(n, 1) = BB_candle_arr(n + 1, 1)
            BB_candle_arr(n, 2) = BB_candle_arr(n + 1, 2)
            BB_candle_arr(n, 3) = BB_candle_arr(n + 1, 3)
        Next

        BB_candle_arr(Form1.candle_init_count - 1, 0) = open_price
        BB_candle_arr(Form1.candle_init_count - 1, 1) = high_price
        BB_candle_arr(Form1.candle_init_count - 1, 2) = low_price
        BB_candle_arr(Form1.candle_init_count - 1, 3) = last_price

        ' Middle band
        middle_band = calculate_mean(BB_candle_arr)

        ' Upper band
        upper_band = middle_band + (calculate_std(middle_band, BB_candle_arr) * Form1.tbx_std_dev)

        ' Lower band
        lower_band = middle_band - (calculate_std(middle_band, BB_candle_arr) * Form1.tbx_std_dev)

        ' Band span
        Upper_Lower_Band_Span = upper_band - lower_band

        ' Previous Candle High and Close above Upper Band
        Prev_Candle_High_and_Close_above_U_Band = BB_candle_arr(Form1.candle_init_count - 2, 3) > upper_band

        ' Previous Candle Low and Close below Lower Band
        Prev_Candle_Low_and_Close_below_L_Band = BB_candle_arr(Form1.candle_init_count - 2, 3) > lower_band

        ' Previous low
        Prev_low = BB_candle_arr(Form1.candle_init_count - 2, 2)

        ' Previous high
        Prev_high = BB_candle_arr(Form1.candle_init_count - 2, 1)

        If Properties_Class.position_opened = True Then

            ' Check for stop loss trigger

            ' Check for take profit trigger

        End If

        If Properties_Class.position_opened = False Then
            ' Check entry trigger
            Call position_entry_test(Upper_Lower_Band_Span, upper_band, Prev_Candle_High_and_Close_above_U_Band, Prev_Candle_Low_and_Close_below_L_Band, high_price, low_price, last_price, Prev_low, Prev_high)

        End If

#End Region


    End Sub

    Function position_entry_test(Upper_Lower_Band_Span As Double, upper_band As Double, Prev_Candle_High_and_Close_above_U_Band As Boolean, Prev_Candle_Low_and_Close_below_L_Band As Boolean, high_price As Double, low_price As Double, last_price As Double, Prev_low As Double, Prev_high As Double)
        ' Long entry check
        If Upper_Lower_Band_Span > 0.0015 And (((low_price < lower_band) And (last_price > lower_band)) Or ((Prev_Candle_Low_and_Close_below_L_Band = True) And (low_price > lower_band))) Then
            ' Assign trade opened flag
            Properties_Class.long_position_opened = True

            ' Assign stop price
            If Prev_Candle_Low_and_Close_below_L_Band = True Then
                If Prev_low < low_price Then
                    Properties_Class.stop_price = Prev_low - 0.0005
                Else
                    Properties_Class.stop_price = low_price - 0.0005
                End If
            Else
                Properties_Class.stop_price = low_price - 0.0005
            End If

            ' Execute long entry trade

        End If

        ' Short entry check
        If Upper_Lower_Band_Span > 0.0015 And (((high_price > upper_band) And (last_price < upper_band)) Or ((Prev_Candle_High_and_Close_above_U_Band = True) And (low_price < upper_band))) Then
            ' Assign trade opened flag
            Properties_Class.short_position_opened = True

            ' Assign stop price
            If Prev_Candle_High_and_Close_above_U_Band = True Then
                If Prev_high > high_price Then
                    Properties_Class.stop_price = Prev_high + 0.0005
                Else
                    Properties_Class.stop_price = high_price + 0.0005
                End If
            Else
                Properties_Class.stop_price = high_price + 0.0005
            End If

            ' Execute short entry trade

        End If

    End Function

    Function stop_loss_check(last_price As Double)
        ' Long position check
        If Properties_Class.long_position_opened = True Then
            If last_price < Properties_Class.stop_price Then
                ' Execute take stop trade

            End If
            ' Short position check
        ElseIf Properties_Class.short_position_opened = True Then
            If last_price > Properties_Class.stop_price Then
                ' Execute take stop trade

            End If
        End If

    End Function

    Function take_profit_check(last_price As Double, middle_band As Double)
        ' Long position check
        If last_price > middle_band + 0.001 And Properties_Class.long_position_opened = True Then
            ' Execute take proift trade

        End If
        ' Short position check
        If last_price < middle_band - 0.001 And Properties_Class.short_position_opened = True Then
            ' Execute take proift trade

        End If

    End Function

    Function calculate_mean(BB_candle_arr(,) As Double)
        Dim sum_of_close As Double

        For n = 0 To Form1.candle_init_count - 1
            sum_of_close = sum_of_close + BB_candle_arr(n, 3)
        Next

        middle_band = sum_of_close / BB_candle_arr.Length

        Return middle_band

    End Function

    Function calculate_std(middle_band As Double, BB_candle_arr(,) As Double)

        Dim value_and_mean_diff As Double
        Dim std As Double

        For n = 0 To Form1.candle_init_count - 1
            value_and_mean_diff = value_and_mean_diff + ((BB_candle_arr(n, 3) - middle_band) * (BB_candle_arr(n, 3) - middle_band))
        Next

        std = Math.Sqrt(value_and_mean_diff / BB_candle_arr.Length)

        Return std

    End Function
End Class
