Imports System.Threading

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

        'For Testing:
        If Raising_Orders.test_completed = True Then

            MsgBox("Testing completed - trade roundtrip completed")

        End If

        ' Check for stop loss trigger
        If Properties_Class.position_opened = True Then
            If Properties_Class.stp_loss_chck_in_progress = True Then
            Else
                Properties_Class.stp_loss_chck_in_progress = True
                Dim stp_loss_check_thrd As Thread

                stp_loss_check_thrd = New Thread(AddressOf stop_loss_check)
                stp_loss_check_thrd.Start()
                stp_loss_check_thrd.Join()
                'Call stop_loss_check(last_price)
                Properties_Class.stp_loss_chck_in_progress = False
            End If
        End If

        ' Check for take profit trigger
        If Properties_Class.position_opened = True Then
            If Properties_Class.take_prft_chck_in_progress = True Then
            Else
                Properties_Class.take_prft_chck_in_progress = True
                Dim take_prft_check_thrd As Thread

                take_prft_check_thrd = New Thread(Sub() take_profit_check(last_price, middle_band))
                take_prft_check_thrd.Start()
                take_prft_check_thrd.Join()
                'Call take_profit_check(last_price, middle_band)
                Properties_Class.take_prft_chck_in_progress = False
            End If
        End If

        ' Check entry trigger
        Do While Properties_Class.position_entry_in_progress = True
            Exit Sub
        Loop
        If Properties_Class.position_opened = False Then
            Call position_entry_test(Upper_Lower_Band_Span, upper_band, Prev_Candle_High_and_Close_above_U_Band, Prev_Candle_Low_and_Close_below_L_Band, high_price, low_price, last_price, Prev_low, Prev_high)

        End If

#End Region


    End Sub

    Function position_entry_test(Upper_Lower_Band_Span As Double, upper_band As Double, Prev_Candle_High_and_Close_above_U_Band As Boolean, Prev_Candle_Low_and_Close_below_L_Band As Boolean, high_price As Double, low_price As Double, last_price As Double, Prev_low As Double, Prev_high As Double)
        ' Long entry check
        Dim i_execute As Execute = New Execute
        Dim i_Auto_open_trade_parameters As Auto_open_trade_parameters = New Auto_open_trade_parameters

        If Upper_Lower_Band_Span > 0.0004 And (((low_price < lower_band) And (last_price > lower_band)) Or ((Prev_Candle_Low_and_Close_below_L_Band = True) And (last_price > lower_band))) Then
            ' Assign trade opened flag
            Properties_Class.position_entry_in_progress = True
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
            Dim exec_thrd As Thread

            exec_thrd = New Thread(AddressOf i_execute.execute_in_ord_LMT)
            i_Auto_open_trade_parameters.Auto_open_trade_action = "BUY"
            i_Auto_open_trade_parameters.Auto_open_trade_price = last_price
            exec_thrd.Start(i_Auto_open_trade_parameters)
            exec_thrd.Join()
        End If

        ' Short entry check
        If Upper_Lower_Band_Span > 0.0015 And (((high_price > upper_band) And (last_price < upper_band)) Or ((Prev_Candle_High_and_Close_above_U_Band = True) And (last_price < upper_band))) Then
            ' Assign trade opened flag
            Properties_Class.position_entry_in_progress = True
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
            Dim exec_thrd As Thread

            exec_thrd = New Thread(AddressOf i_execute.execute_in_ord_LMT)
            i_Auto_open_trade_parameters.Auto_open_trade_action = "SELL"
            i_Auto_open_trade_parameters.Auto_open_trade_price = last_price
            exec_thrd.Start(i_Auto_open_trade_parameters)
            exec_thrd.Join()
        End If

    End Function

    Function stop_loss_check(last_price As Double)
        Dim i_execute As Execute = New Execute
        Dim i_Auto_open_trade_parameters As Auto_open_trade_parameters = New Auto_open_trade_parameters

        ' Long position check
        If Properties_Class.long_position_opened = True Then
            If last_price < Properties_Class.stop_price Then
                Properties_Class.long_position_opened = False
                ' Execute take stop trade
                Dim stop_loss_exec_thrd As Thread

                stop_loss_exec_thrd = New Thread(AddressOf i_execute.execute_in_ord_LMT)
                i_Auto_open_trade_parameters.Auto_open_trade_action = "SELL"
                i_Auto_open_trade_parameters.Auto_open_trade_price = last_price
                stop_loss_exec_thrd.Start(i_Auto_open_trade_parameters)
                stop_loss_exec_thrd.Join()
            End If
            ' Short position check
        ElseIf Properties_Class.short_position_opened = True Then
            If last_price > Properties_Class.stop_price Then
                Properties_Class.short_position_opened = False
                ' Execute take stop trade
                Dim stop_loss_exec_thrd As Thread

                stop_loss_exec_thrd = New Thread(AddressOf i_execute.execute_in_ord_LMT)
                i_Auto_open_trade_parameters.Auto_open_trade_action = "BUY"
                i_Auto_open_trade_parameters.Auto_open_trade_price = last_price
                stop_loss_exec_thrd.Start(i_Auto_open_trade_parameters)
                stop_loss_exec_thrd.Join()
            End If
        End If

    End Function

    Function take_profit_check(last_price As Double, middle_band As Double)
        Dim i_execute As Execute = New Execute
        Dim i_Auto_open_trade_parameters As Auto_open_trade_parameters = New Auto_open_trade_parameters

        ' Long position check
        If last_price > middle_band + 0.001 And Properties_Class.long_position_opened = True Then
            Properties_Class.long_position_opened = False
            ' Execute take proift trade
            Dim take_profit_exec_thrd As Thread

            take_profit_exec_thrd = New Thread(AddressOf i_execute.execute_in_ord_LMT)
            i_Auto_open_trade_parameters.Auto_open_trade_action = "SELL"
            i_Auto_open_trade_parameters.Auto_open_trade_price = last_price
            take_profit_exec_thrd.Start(i_Auto_open_trade_parameters)
            take_profit_exec_thrd.Join()
        End If
        ' Short position check
        If last_price < middle_band - 0.001 And Properties_Class.short_position_opened = True Then
            Properties_Class.short_position_opened = False
            ' Execute take proift trade
            Dim take_profit_exec_thrd As Thread

            take_profit_exec_thrd = New Thread(AddressOf i_execute.execute_in_ord_LMT)
            i_Auto_open_trade_parameters.Auto_open_trade_action = "BUY"
            i_Auto_open_trade_parameters.Auto_open_trade_price = last_price
            take_profit_exec_thrd.Start(i_Auto_open_trade_parameters)
            take_profit_exec_thrd.Join()
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
