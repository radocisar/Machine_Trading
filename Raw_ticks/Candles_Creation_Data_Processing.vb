Imports System.IO
Imports System.Threading

Public Class Candles_Creation_Data_Processing

    Private BID_price As Double
    Private ASK_price As Double
    Private Current_time As DateTime
    Dim mls_200 As TimeSpan
    Dim i_candle As Candle = New Candle
    Private candle_arr() As Candle
    Public candle_resolution As Integer = CInt(Form1.lbl_candle_resolution.Text)
    Public candle_init_count As Integer = CInt(Form1.lbl_min_num_of_candles.Text)
    Public Shared Event start_trading_strategy(ByVal candle_arr() As Candle, ByVal open_price As Double, ByVal high_price As Double, ByVal low_price As Double, ByVal last_price As Double)
    Public live_updates_started As Boolean

    Sub mm_price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer)

        Dim open_price As Double
        Dim high_price As Double
        Dim low_price As Double
        Dim last_price As Double
        Dim current_price As Double
        Dim first_top_of_minute_passed As Boolean
        Dim not_first_tick_of_minute As Boolean
        Dim prev_tick_second As Integer
        Dim minute_candle_non_zero_sec_ticking_in_progress As Boolean
        Dim first_candle_start_time As DateTime
        Dim first_candle_end_time As DateTime
        Dim not_first_tick As Boolean
        Dim candle_start_time As DateTime
        Dim candle_end_time As DateTime
        Dim first_tick_of_first_candle As Boolean
        Dim live_updates_started As Boolean

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
                current_price = price
                Exit Sub
        End Select

#Region "Filling in minutely candles"

        Current_time = DateTime.Now

        ' Before first candle starts to be filled
        If first_top_of_minute_passed = False Then
            If not_first_tick = False Then
                first_candle_start_time = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, Current_time.Hour, Current_time.Minute, 0)
                first_candle_end_time = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, Current_time.Hour, Current_time.Minute, 59)
                not_first_tick = True
            End If
            If Current_time >= first_candle_start_time + New TimeSpan(0, 0, 60) Then
                first_top_of_minute_passed = True
                candle_start_time = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, Current_time.Hour, Current_time.Minute, 0)
                candle_end_time = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, Current_time.Hour, Current_time.Minute, 59)
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
            candle_start_time = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, Current_time.Hour, Current_time.Minute, 0)
            candle_end_time = New DateTime(Current_time.Year, Current_time.Month, Current_time.Day, Current_time.Hour, Current_time.Minute, 59)
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
#End Region



    End Sub

#Region "Creating candles array"
    Sub candle_array(open_price As Double, high_price As Double, low_price As Double, last_price As Double, origin As String)

        Select Case origin
            Case "new-candle"
                ReDim Preserve candle_arr(candle_init_count - 1)

                For n = 0 To candle_init_count - 2
                    candle_arr(n).open = candle_arr(n + 1).open
                    candle_arr(n).high = candle_arr(n + 1).high
                    candle_arr(n).low = candle_arr(n + 1).low
                    candle_arr(n).close = candle_arr(n + 1).close
                Next

                candle_arr(candle_init_count - 1).open = open_price
                candle_arr(candle_init_count - 1).high = high_price
                candle_arr(candle_init_count - 1).low = low_price
                candle_arr(candle_init_count - 1).close = last_price

                If candle_arr(0).close <> "" Then
                    ' start trading the strategy
                    live_updates_started = True
                    'RaiseEvent start_trading_strategy(candle_arr)
                End If
            Case "intra-candle"
                RaiseEvent start_trading_strategy(candle_arr, open_price, high_price, low_price, last_price)
        End Select


    End Sub
#End Region

End Class
