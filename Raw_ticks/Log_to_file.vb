Imports System.IO

Public Class Log_to_file
    Public Shared main_lbl As String
    Public Shared i_ordr_Action_lbl As String
    Public Shared price_lbl As String
    Public Shared i_ordr_Quantity_lbl As String
    Public Shared i_ordr_OrderType_lbl As String
    Public Shared orderId_lbl As String
    Public Shared status_lbl As String
    Public Shared ma_5_ticks_lbl As String
    Public Shared ma_5_ticks_prior_per_lbl As String
    Public Shared ma_10_ticks_lbl As String
    Public Shared ma_10_ticks_prior_per_lbl As String
    Public Shared ma_5_ticks_for_exit_calc_lbl As String
    Public Shared ma_10_ticks_for_exit_calc_lbl As String
    Public Shared thread_lock As New Object
    Public Shared thread_lock_for_candles_logging As New Object
    Public Shared thread_lock_for_execution_logging As New Object


    Public Shared Sub Log_to_file_method(Action As String, price As Double, Quantity As Integer, OrderType As String, orderId As Integer, status As String, ma_5_ticks As Double, ma_5_ticks_prior_per As Double, ma_10_ticks As Double,
                                         ma_10_ticks_prior_per As Double, ma_5_ticks_for_exit_calc As Double, ma_10_ticks_for_exit_calc As Double)

        'Dim str_wrt As StreamWriter = New StreamWriter("C:\Raw_Data\Logging _Auto_Trading\log.txt", True)
        SyncLock thread_lock
            Open_str_wrt.str_wrt.WriteLine(main_lbl & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & i_ordr_Action_lbl & Action & vbNewLine & price_lbl & price & vbNewLine & i_ordr_Quantity_lbl _
            & Quantity & vbNewLine & i_ordr_OrderType_lbl & OrderType & vbNewLine & orderId_lbl & orderId & vbNewLine & status_lbl & status & vbNewLine & ma_5_ticks_lbl & ma_5_ticks & vbNewLine & ma_5_ticks_prior_per_lbl & ma_5_ticks_prior_per &
            vbNewLine & ma_10_ticks_lbl & ma_10_ticks & vbNewLine & ma_10_ticks_prior_per_lbl & ma_10_ticks_prior_per & vbNewLine & vbNewLine & ma_5_ticks_for_exit_calc_lbl & ma_5_ticks_for_exit_calc & vbNewLine &
            ma_10_ticks_for_exit_calc_lbl & ma_10_ticks_for_exit_calc & vbNewLine)
        End SyncLock
        'str_wrt.Close()

    End Sub

    Public Shared Sub Log_to_file_method_for_candles_logging(origin As String, Current_time As DateTime, candle_start_time As DateTime, candle_end_time As DateTime, date_span_to_subtract_for_logging As TimeSpan, candle_arr(,) As Double,
                                                             open_price As Double, high_price As Double, low_price As Double, last_price As Double)

        SyncLock thread_lock_for_candles_logging
            Select Case origin
                Case "new-candle"
                    Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine(Current_time)
                    Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Start: " & candle_start_time.Subtract(date_span_to_subtract_for_logging) & "|" & "End: " & candle_end_time.Subtract(date_span_to_subtract_for_logging))
                    For n = 0 To Form1.candle_init_count - 1
                        Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle " & n + 1 & " - O: " & candle_arr(n, 0))
                        Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle " & n + 1 & " - H: " & candle_arr(n, 1))
                        Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle " & n + 1 & " - L: " & candle_arr(n, 2))
                        Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle " & n + 1 & " - C: " & candle_arr(n, 3))
                    Next
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 1 - O: " & candle_arr(0, 0))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 1 - H: " & candle_arr(0, 1))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 1 - L: " & candle_arr(0, 2))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 1 - C: " & candle_arr(0, 3))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 2 - O: " & candle_arr(1, 0))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 2 - H: " & candle_arr(1, 1))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 2 - L: " & candle_arr(1, 2))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 2 - C: " & candle_arr(1, 3))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 3 - O: " & candle_arr(2, 0))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 3 - H: " & candle_arr(2, 1))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 3 - L: " & candle_arr(2, 2))
                'Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Candle 3 - C: " & candle_arr(2, 3))
                Case "intra-candle"
                    Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine(Current_time)
                    Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine(candle_start_time & "|" & candle_end_time)
                    Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine(open_price & "|" & high_price & "|" & low_price & "|" & last_price)
                    Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("     ")
            End Select
        End SyncLock
    End Sub

    Public Shared Sub Log_to_file_method_for_execution(Log_Upper_Lower_Band_Span As Double, Log_upper_band As Double, Log_middle_band As Double, Log_lower_band As Double, Log_Prev_Candle_High_and_Close_above_U_Band As Double, Log_Prev_Candle_Low_and_Close_below_L_Band As Double, Log_high_price As Double,
                                                       Log_low_price As Double, Log_last_price As Double, Log_Prev_low As Double, Log_Prev_high As Double)

        SyncLock thread_lock_for_execution_logging
            Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine(DateTime.Now)
            Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("Execution Details: " & vbCrLf &
                                                               "Upper_Lower_Band_Span: " & Log_Upper_Lower_Band_Span & vbCrLf &
                                                               "Upper_band: " & Log_upper_band & vbCrLf &
                                                               "Middle_band" & Log_middle_band & vbCrLf &
                                                               "Lower_band" & Log_lower_band & vbCrLf &
                                                               "Prev_Candle_High_and_Close_above_U_Band: " & Log_Prev_Candle_High_and_Close_above_U_Band & vbCrLf &
                                                               "Prev_Candle_Low_and_Close_below_L_Band: " & Log_Prev_Candle_Low_and_Close_below_L_Band & vbCrLf &
                                                               "high_price: " & Log_high_price & vbCrLf &
                                                               "low_price: " & Log_low_price & vbCrLf &
                                                               "last_price: " & Log_last_price & vbCrLf &
                                                               "Prev_low: " & Log_Prev_low & vbCrLf &
                                                               "Prev_high: " & Log_Prev_high)
            Open_str_wrt.str_wrt_for_Candles_Logging.WriteLine("     ")
        End SyncLock
    End Sub
End Class
