Imports System.IO
Imports System.Threading

Public Class Candles_Creation_Data_Processing

    Private BID_price As Double
    Private ASK_price As Double
    Private Current_time As DateTime
    Dim mls_200 As TimeSpan
    Dim i_candle As Candle = New Candle
    Private candle_arr(Form1.candle_init_count - 1, 3) As Double
    'Public candle_resolution As Integer = CInt(Form1.lbl_candle_resolution.Text)
    'Public candle_init_count As Integer = CInt(Form1.lbl_min_num_of_candles.Text)
    Public Shared Event start_trading_strategy(ByVal candle_arr() As Double, ByVal open_price As Double, ByVal high_price As Double, ByVal low_price As Double, ByVal last_price As Double)
    Public live_updates_started As Boolean
    Public not_first_tick As Boolean
    Public first_top_of_minute_passed As Boolean
    Public first_candle_start_time As DateTime
    Public first_candle_end_time As DateTime
    Public candle_start_time As DateTime
    Public candle_end_time As DateTime
    Public prev_candle_start_time As DateTime
    Public open_price As Double
    Public high_price As Double
    Public low_price As Double
    Public last_price As Double
    Public candle_update_open_price As Double
    Public candle_update_high_price As Double
    Public candle_update_low_price As Double
    Public candle_update_last_price As Double
    Public current_price As Double
    Private start_candles_population As Boolean
    Private candle_array_update_lock As Boolean
    Public pause_live_ticks As Boolean


    'Dim sw As StreamWriter

    Sub mm_price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer)


        'Dim not_first_tick_of_minute As Boolean
        'Dim prev_tick_second As Integer
        'Dim minute_candle_non_zero_sec_ticking_in_progress As Boolean
        'Dim first_tick_of_first_candle As Boolean
        Dim thrd_candle_creation As Thread

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
                start_candles_population = True
                prev_candle_start_time = Calculate_start_time_of_candle(Current_time) ' Setting the previous candle start time to the same time as current candle's start time as this is a one-off operation here. Afterwards, this operation 
                ' will happen in the candle_array() function only when a candle is finished creating
                'candle_start_time = Calculate_start_time_of_candle(Current_time)
                'candle_end_time = Calculate_end_time_of_candle(Current_time)
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
        If start_candles_population = True Then
            candle_start_time = Calculate_start_time_of_candle(Current_time)
            candle_end_time = Calculate_end_time_of_candle(Current_time)

            If candle_start_time > prev_candle_start_time Then ' Ensures that when a new candle time starts the last completed candle is written to the candle_array()
                If candle_array_update_lock = True Then ' In conjunction with "candle_start_time > prev_candle_start_time" ensures that while candle_update_* prices are being updated with the completed candle's last values no other tick comes through 
                    'to update them again while another tick may have updated open_price, high_price, low_price and last_price meanwhile. All such ticks are discarted (there is very likely to be very few of these)
                    Do While candle_start_time > prev_candle_start_time
                        Exit Sub
                    Loop
                End If

                candle_array_update_lock = True

                candle_update_open_price = open_price
                candle_update_high_price = high_price
                candle_update_low_price = low_price
                candle_update_last_price = last_price

                open_price = current_price
                high_price = current_price
                low_price = current_price
                last_price = current_price

                prev_candle_start_time = candle_start_time

                thrd_candle_creation = New Thread(Sub() candle_array(candle_update_open_price, candle_update_high_price, candle_update_low_price, candle_update_last_price, "new-candle", Current_time))
                thrd_candle_creation.Start()
                'Call candle_array(candle_update_open_price, candle_update_high_price, candle_update_low_price, candle_update_last_price, "new-candle", Current_time)

                'volume =
                If live_updates_started = True Then
                    Call candle_array(open_price, high_price, low_price, last_price, "intra-candle", Current_time)
                End If
                candle_array_update_lock = False
                Exit Sub
            End If

            'If (candle_start_time <= Current_time) And (Current_time <= candle_end_time) Then
            Do While pause_live_ticks = True
                Continue Do
            Loop
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
                Call candle_array(open_price, high_price, low_price, last_price, "intra-candle", Current_time)
            End If
            'Else ' next candle
            ' assigning candle
            'candle_start_time = Calculate_start_time_of_candle(Current_time)
            'candle_end_time = Calculate_end_time_of_candle(Current_time)

            'candle_start_time = Calculate_start_time_of_candle(Current_time)
            'candle_end_time = Calculate_end_time_of_candle(Current_time)

            'End If

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
    Sub candle_array(candle_update_open_price As Double, candle_update_high_price As Double, candle_update_low_price As Double, candle_update_last_price As Double, origin As String, Current_time As DateTime)

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

                candle_arr(Form1.candle_init_count - 1, 0) = candle_update_open_price
                candle_arr(Form1.candle_init_count - 1, 1) = candle_update_high_price
                candle_arr(Form1.candle_init_count - 1, 2) = candle_update_low_price
                candle_arr(Form1.candle_init_count - 1, 3) = candle_update_last_price

                If candle_arr(0, 0) <> Nothing Then
                    ' pause other live ticks which may have arrived outside of the candle creation if-statement block
                    pause_live_ticks = True
                    ' start trading the strategy
                    live_updates_started = True
                    'RaiseEvent start_trading_strategy(candle_arr, open_price, high_price, low_price, last_price)
                    Open_str_wrt.str_wrt_1.WriteLine(Current_time)
                    Open_str_wrt.str_wrt_1.WriteLine("Start: " & candle_start_time.AddSeconds(120) & "|" & "End: " & candle_end_time.AddSeconds(239))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 1 - O: " & candle_arr(0, 0))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 1 - H: " & candle_arr(0, 1))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 1 - L: " & candle_arr(0, 2))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 1 - C: " & candle_arr(0, 3))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 2 - O: " & candle_arr(1, 0))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 2 - H: " & candle_arr(1, 1))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 2 - L: " & candle_arr(1, 2))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 2 - C: " & candle_arr(1, 3))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 3 - O: " & candle_arr(2, 0))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 3 - H: " & candle_arr(2, 1))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 3 - L: " & candle_arr(2, 2))
                    Open_str_wrt.str_wrt_1.WriteLine("Candle 3 - C: " & candle_arr(2, 3))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 4 - O: " & candle_arr(3, 0))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 4 - H: " & candle_arr(3, 1))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 4 - L: " & candle_arr(3, 2))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 4 - C: " & candle_arr(3, 3))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 5 - O: " & candle_arr(4, 0))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 5 - H: " & candle_arr(4, 1))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 5 - L: " & candle_arr(4, 2))
                    'Open_str_wrt.str_wrt_1.WriteLine("Candle 5 - C: " & candle_arr(4, 3))
                End If
            Case "intra-candle"
                'RaiseEvent start_trading_strategy(candle_arr, open_price, high_price, low_price, last_price)
                Open_str_wrt.str_wrt_1.WriteLine(Current_time)
                Open_str_wrt.str_wrt_1.WriteLine(candle_start_time & "|" & candle_end_time)
                Open_str_wrt.str_wrt_1.WriteLine(open_price & "|" & high_price & "|" & low_price & "|" & last_price)
                Open_str_wrt.str_wrt_1.WriteLine("     ")
                pause_live_ticks = False
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
