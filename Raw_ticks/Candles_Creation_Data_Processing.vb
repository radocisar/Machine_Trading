Imports System.IO
Imports System.Threading

Public Class Candles_Creation_Data_Processing

    Private BID_price As Double
    Private ASK_price As Double
    Private Current_time As DateTime
    Dim mls_200 As TimeSpan
    Dim i_candle As Candle = New Candle
    Private candle_arr(4, 3) As Double
    'Public candle_resolution As Integer = CInt(Form1.lbl_candle_resolution.Text)
    'Public candle_init_count As Integer = CInt(Form1.lbl_min_num_of_candles.Text)
    Public Shared Event start_trading_strategy(ByVal candle_arr() As Candle, ByVal open_price As Double, ByVal high_price As Double, ByVal low_price As Double, ByVal last_price As Double)
    Public live_updates_started As Boolean
    Public not_first_tick As Boolean
    Public first_top_of_minute_passed As Boolean
    Public first_candle_start_time As DateTime
    Public first_candle_end_time As DateTime
    Public candle_start_time As DateTime
    Public candle_end_time As DateTime
    Public open_price As Double
    Public high_price As Double
    Public low_price As Double
    Public last_price As Double
    Public current_price As Double

    'Dim sw As StreamWriter

    Sub mm_price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer)


        'Dim not_first_tick_of_minute As Boolean
        'Dim prev_tick_second As Integer
        'Dim minute_candle_non_zero_sec_ticking_in_progress As Boolean
        'Dim first_tick_of_first_candle As Boolean

        Select Case field
            Case 1
                'tick_type = "BID"
                BID_price = price
            Case 2
                'tick_type = "ASK"
                ASK_price = price
            Case 6
                'tick_type = "HIGH"
                Exit Sub
            Case 7
                'tick_type = "LOW"
                Exit Sub
            Case 4
                'tick_type = "LAST"
                'current_price = price
                Exit Sub
        End Select

        If BID_price <> 0 And ASK_price <> 0 Then
            current_price = (BID_price + ASK_price) / 2
        Else
            Exit Sub
        End If

#Region "Filling in minutely candles"

        Current_time = DateTime.Now

        ' Before first candle starts to be filled
        If first_top_of_minute_passed = False Then
            If not_first_tick = False Then

                first_candle_start_time = Calculate_start_time_of_candle(Current_time)
                first_candle_end_time = Calculate_end_time_of_candle(Current_time)
                not_first_tick = True
            End If
            If Current_time >= first_candle_start_time + New TimeSpan(0, 0, 60) Then
                first_top_of_minute_passed = True
                candle_start_time = Calculate_start_time_of_candle(Current_time)
                candle_end_time = Calculate_end_time_of_candle(Current_time)
                'first_tick_of_first_candle = True
                open_price = current_price
                high_price = current_price
                low_price = current_price
                last_price = current_price
                'volume =
            Else
                Exit Sub
            End If
        End If

        ' After first candle starts to be filled
        If (candle_start_time <= Current_time) And (Current_time <= candle_end_time) Then
            ' Assign last price
            last_price = current_price
            ' Assign high price
            If current_price > high_price Then
                high_price = current_price
            End If
            ' Assign low price
            If current_price < low_price Then
                low_price = current_price
            End If
            ' Assign volume

            If live_updates_started = True Then
                Call candle_array(open_price, high_price, low_price, last_price, "intra-candle")
            End If
        Else ' next candle
            ' assigning candle
            Call candle_array(open_price, high_price, low_price, last_price, "new-candle")
            candle_start_time = Calculate_start_time_of_candle(Current_time)
            candle_end_time = Calculate_end_time_of_candle(Current_time)
            open_price = current_price
            high_price = current_price
            low_price = current_price
            last_price = current_price
            'volume =
            If live_updates_started = True Then
                Call candle_array(open_price, high_price, low_price, last_price, "intra-candle")
            End If

        End If

        'If Current_time.Second <> 0 And not_first_tick_of_minute = True Then
        '    minute_candle_non_zero_sec_ticking_in_progress = True
        'End If

        'If Current_time.Second = 0 And minute_candle_non_zero_sec_ticking_in_progress = True Then
        '    not_first_tick_of_minute = False
        'End If

        ' After first candle starts to be filled
        'If Current_time.Second <> 0 Or (Current_time.Second = 0 And not_first_tick_of_minute = True) Then


        'ElseIf Current_time.Second = 0 And not_first_tick_of_minute = False And last_price <> "" Then
        '    assign candle into the array
        '        open_price = current_price
        '    high_price = current_price
        '    low_price = current_price
        '    last_price = current_price
        '    not_first_tick_of_minute = True
        'End If

        'mls_200 = New TimeSpan(0, 0, 0, 0, 200)

    End Sub
#End Region

#Region "Creating candles array"
    Sub candle_array(open_price As Double, high_price As Double, low_price As Double, last_price As Double, origin As String)

        Select Case origin
            Case "new-candle"
                ReDim Preserve candle_arr(Form1.candle_init_count - 1, 3)

                For n = 0 To Form1.candle_init_count - 2
                    ' First shift (when only last slot of the array is populated) and ensures that "open", "high", "low" and "close" are created for candle_arr()
                    If candle_arr(Form1.candle_init_count - 1, 0) <> Nothing Then
                        candle_arr(n, 0) = candle_arr(n + 1, 0)
                        candle_arr(n, 1) = candle_arr(n + 1, 1)
                        candle_arr(n, 2) = candle_arr(n + 1, 2)
                        candle_arr(n, 3) = candle_arr(n + 1, 3)
                    End If
                Next

                candle_arr(Form1.candle_init_count - 1, 0) = open_price
                candle_arr(Form1.candle_init_count - 1, 1) = high_price
                candle_arr(Form1.candle_init_count - 1, 2) = low_price
                candle_arr(Form1.candle_init_count - 1, 3) = last_price

                If candle_arr(0, 0) <> Nothing Then
                    ' start trading the strategy
                    live_updates_started = True
                    'RaiseEvent start_trading_strategy(candle_arr)
                    Open_str_wrt.str_wrt_1.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff"))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(0, 0))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(0, 1))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(0, 2))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(0, 3))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(1, 0))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(1, 1))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(1, 2))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(1, 3))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(2, 0))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(2, 1))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(2, 2))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(2, 3))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(3, 0))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(3, 1))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(3, 2))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(3, 3))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(4, 0))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(4, 1))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(4, 2))
                    Open_str_wrt.str_wrt_1.WriteLine(candle_arr(4, 3))
                End If
            Case "intra-candle"
                'RaiseEvent start_trading_strategy(candle_arr, open_price, high_price, low_price, last_price)
                Open_str_wrt.str_wrt_1.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff"))
                Open_str_wrt.str_wrt_1.WriteLine(open_price, high_price, low_price, last_price)
                Open_str_wrt.str_wrt_1.WriteLine("     ")
        End Select


    End Sub
#End Region

#Region "Functions"
    Function Calculate_start_time_of_candle(Current_time As DateTime)

        Dim current_midnight As DateTime

        current_midnight = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, 0, 0, 0)

        For n = 0 To ((1440 / Form1.candle_resolution) - 1)
            If (current_midnight.AddMinutes(n * Form1.candle_resolution) < Current_time) And (Current_time < current_midnight.AddMinutes((n + 1) * Form1.candle_resolution)) Then
                Return current_midnight.AddMinutes(n * Form1.candle_resolution)
            End If
        Next

    End Function

    Function Calculate_end_time_of_candle(Current_time As DateTime)

        Dim current_midnight As DateTime
        Dim intermediate_end_time As DateTime


        current_midnight = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, 0, 0, 0)

        For n = 0 To ((1440 / Form1.candle_resolution) - 1)
            intermediate_end_time = current_midnight.AddMinutes((n + 1) * Form1.candle_resolution - 1)
            If (current_midnight.AddMinutes(n * Form1.candle_resolution) < Current_time) And (Current_time < intermediate_end_time.AddSeconds(59)) Then
                Return intermediate_end_time.AddSeconds(59)
            End If
        Next

    End Function
#End Region
End Class
