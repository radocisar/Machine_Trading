Imports System.IO
Imports System.Threading

Public Class Data_Requests_Handlers_ma_5_10

    'Dim dta As Data = New Data
    Dim i_Conditions_analysis_ma_5_10 As Conditions_analysis_ma_5_10 = New Conditions_analysis_ma_5_10
    Dim i_Open_positions_cond_analysis_ma_5_10 As Open_positions_cond_analysis_ma_5_10 = New Open_positions_cond_analysis_ma_5_10
    Dim i_Raising_Orders As Raising_Orders = New Raising_Orders
    Public ma_5_ticks As Double
    'Dim ma_3_ticks As Double
    Public ma_10_ticks As Double
    Public ma_mixed_for_long_entry As Boolean
    Public very_first_time As Boolean
    Public end_very_first_time As Boolean
    Public ma_mixed_for_short_entry As Boolean
    Private ma_ticks_arr() As Double
    Private BID_price As Double
    Private ASK_price As Double
    'Public ma_5_10_position_opened As Boolean

    'For testing
    Public Shared first_trade As Boolean
    Public Shared first_exit_trade As Boolean
    Public Shared tst_data_time As DateTime
    Dim two_seconds As TimeSpan

    Sub mm_string_return_handler(tickerId As Integer, field As Integer, string_return As String)

        'Dim tick_type As String
        'Dim dt As DateTime
        Dim VOLUME_TIME_SALES_arr() As String
        'Dim US_tm_6_hrs As TimeSpan
        'Dim ma_ticks_arr() As Double
        'Dim ma_3_ticks_arr() As Double
        Dim ma_5_ticks_prior_per As Double
        Dim ma_10_ticks_prior_per As Double
        Dim thrd_open_pos_cond_analysis As Thread
        Dim LAST_price As Double

        'ReDim ma_5_ticks_arr(5)
        'ReDim ma_3_ticks_arr(3)
        ReDim Preserve ma_ticks_arr(10)

        Select Case field
            'Case 32
            '    'tick_type = "BID_EXCH"
            '    tick_type = "B_E"
            '    sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & string_return)
            'Case 33
            '    'tick_type = "ASK_EXCH"
            '    tick_type = "A_E"

            '    sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & string_return)
            'Case 45
            '    'tick_type = "LAST_TIMESTAMP"
            '    tick_type = "L_T"

            '    dt = New DateTime(1970, 1, 1, 0, 0, 0)
            '    dt = dt.AddSeconds(CInt(string_return))

            '    If dta.bln = True Then
            '        sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & dt)
            '        sw.WriteLine("----------------------------------------------------")

            '        dta.bln = False

            '    Else
            '        sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & dt)
            '    End If

            ''Case 47
            ''    tick_type = "FUND_RATIOS"
            ''    sw.WriteLine(tick_type & ": " & string_return)
            Case 48
                'tick_type = "VOLUME_TIME_SALES"

                'two_seconds = New TimeSpan(0, 0, 2)

                'If Now > tst_data_time + two_seconds Then
                VOLUME_TIME_SALES_arr = string_return.Split(";")
                LAST_price = VOLUME_TIME_SALES_arr(0)
                'tst_data_time = Now
                'Else
                'Exit Select
                'End If

                'dt = New DateTime(1970, 1, 1, 0, 0, 0)

                'dt = dt.AddMilliseconds(CLng(VOLUME_TIME_SALES_arr(2)))
                'US_tm_6_hrs = New TimeSpan(5, 0, 0)

                'sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & "T_S|" & (dt - US_tm_6_hrs).ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & LAST_price & "|" & VOLUME_TIME_SALES_arr(1) & _
                '            "|" & VOLUME_TIME_SALES_arr(4) & "|" & VOLUME_TIME_SALES_arr(3) * 100 & "|")

                'ma_5_ticks_arr(0) = ma_5_ticks_arr(1)
                'ma_5_ticks_arr(1) = ma_5_ticks_arr(2)
                'ma_5_ticks_arr(2) = ma_5_ticks_arr(3)
                'ma_5_ticks_arr(3) = ma_5_ticks_arr(4)
                'ma_5_ticks_arr(4) = ma_5_ticks_arr(5)
                'ma_5_ticks_arr(5) = LAST_price

                'ma_3_ticks_arr(0) = ma_3_ticks_arr(1)
                'ma_3_ticks_arr(1) = ma_3_ticks_arr(2)
                'ma_3_ticks_arr(2) = LAST_price

                If LAST_price <> "" Then

                    ma_ticks_arr(0) = ma_ticks_arr(1)
                    ma_ticks_arr(1) = ma_ticks_arr(2)
                    ma_ticks_arr(2) = ma_ticks_arr(3)
                    ma_ticks_arr(3) = ma_ticks_arr(4)
                    ma_ticks_arr(4) = ma_ticks_arr(5)
                    ma_ticks_arr(5) = ma_ticks_arr(6)
                    ma_ticks_arr(6) = ma_ticks_arr(7)
                    ma_ticks_arr(7) = ma_ticks_arr(8)
                    ma_ticks_arr(8) = ma_ticks_arr(9)
                    ma_ticks_arr(9) = ma_ticks_arr(10)
                    ma_ticks_arr(10) = LAST_price

                End If

                '-----Starts thread checking for whether opened long positions should be closed
                'for testing
                'i_Raising_Orders.position_opened = False
                'i_Raising_Orders.position_opened
                If Properties_Class.position_opened = True Then
                    i_Open_positions_cond_analysis_ma_5_10.ma_5_ticks_for_exit_calc = ma_5_ticks
                    i_Open_positions_cond_analysis_ma_5_10.ma_10_ticks_for_exit_calc = ma_10_ticks
                    i_Open_positions_cond_analysis_ma_5_10.price_for_exit_calc = ma_ticks_arr(10)
                    thrd_open_pos_cond_analysis = New Thread(AddressOf i_Open_positions_cond_analysis_ma_5_10.open_positions_cond_analysis_ma_5_over_10)
                    'For testing
                    If first_exit_trade = True Then

                    Else
                        thrd_open_pos_cond_analysis.Start(ma_ticks_arr)
                    End If

                End If

                'TODO Only if condition is re-met after a reversal (not every time conditions are met)

                'If ma_ticks_arr(7) <> Nothing And LAST_price <> "" Then
                If ma_ticks_arr(0) <> Nothing And LAST_price <> "" Then

                    'ma_5_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9)) / 2
                    'ma_5_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8)) / 2
                    'ma_10_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8)) / 3
                    'ma_10_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7)) / 3

                    ma_5_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6)) / 5
                    ma_5_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5)) / 5
                    'ma_3_ticks = (ma_3_ticks_arr(2) + ma_3_ticks_arr(1) + ma_3_ticks_arr(0)) / 3
                    ma_10_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5) + ma_ticks_arr(4) + ma_ticks_arr(3) + ma_ticks_arr(2) + ma_ticks_arr(1)) / 10
                    ma_10_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5) + ma_ticks_arr(4) + ma_ticks_arr(3) + ma_ticks_arr(2) + ma_ticks_arr(1) + ma_ticks_arr(0)) / 10


                    '------Ensures that tick is a very first tick, when it is
                    If end_very_first_time = False Then
                        very_first_time = True
                    End If

                    '-----Ensures that very first met condition is ignored, as position should only be entered after condition is met coming from a non-met state
                    If ma_5_ticks < ma_10_ticks And very_first_time = True Then
                        ma_mixed_for_long_entry = True
                        end_very_first_time = True
                        very_first_time = False
                    Else
                        If ma_5_ticks > ma_10_ticks Then
                            ma_mixed_for_short_entry = True
                        End If
                    End If

                    '-----Call conditions analysis calculations for long entry
                    If ma_mixed_for_long_entry = True Then
                        'For testing
                        If first_trade = True Then

                        Else
                            Call i_Conditions_analysis_ma_5_10.conditions_analysis_ma_5_over_10(ma_5_ticks, ma_10_ticks, ma_5_ticks_prior_per, ma_10_ticks_prior_per, LAST_price)
                        End If

                    End If

                    '-----Call conditions analysis calculations for short entry
                    If ma_mixed_for_short_entry = True Then

                        'Call i_Conditions_analysis_ma_5_10.conditions_analysis_ma_5_under_10(ma_5_ticks, ma_10_ticks, ma_5_ticks_prior_per, ma_10_ticks_prior_per, LAST_price)

                    End If

                End If

                Erase VOLUME_TIME_SALES_arr

        End Select

        'cntr_id = tickerId
        'can_auto_exec = canAutoExecute

    End Sub


    Sub mm_price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer)

        'Dim tick_type As String
        'Dim dt As DateTime
        'Dim VOLUME_TIME_SALES_arr() As String
        'Dim US_tm_6_hrs As TimeSpan
        'Dim ma_ticks_arr() As Double
        'Dim ma_3_ticks_arr() As Double
        Dim ma_5_ticks_prior_per As Double
        Dim ma_10_ticks_prior_per As Double
        Dim thrd_open_pos_cond_analysis As Thread
        Dim MID_price As Double

        'ReDim ma_5_ticks_arr(5)
        'ReDim ma_3_ticks_arr(3)
        ReDim Preserve ma_ticks_arr(10)

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
            Case 9
                'tick_type = "CLOSE"
                Exit Sub
        End Select

        'tick_type = "VOLUME_TIME_SALES"

        'VOLUME_TIME_SALES_arr = string_return.Split(";")

        'dt = New DateTime(1970, 1, 1, 0, 0, 0)

        'dt = dt.AddMilliseconds(CLng(VOLUME_TIME_SALES_arr(2)))
        'US_tm_6_hrs = New TimeSpan(5, 0, 0)

        If BID_price <> 0 And ASK_price <> 0 Then

            MID_price = (BID_price + ASK_price) / 2

        End If

        If MID_price <> 0 Then

            ma_ticks_arr(0) = ma_ticks_arr(1)
            ma_ticks_arr(1) = ma_ticks_arr(2)
            ma_ticks_arr(2) = ma_ticks_arr(3)
            ma_ticks_arr(3) = ma_ticks_arr(4)
            ma_ticks_arr(4) = ma_ticks_arr(5)
            ma_ticks_arr(5) = ma_ticks_arr(6)
            ma_ticks_arr(6) = ma_ticks_arr(7)
            ma_ticks_arr(7) = ma_ticks_arr(8)
            ma_ticks_arr(8) = ma_ticks_arr(9)
            ma_ticks_arr(9) = ma_ticks_arr(10)
            ma_ticks_arr(10) = MID_price

        End If

        '-----Starts thread checking for whether opened long positions should be closed
        'for testing
        'i_Raising_Orders.position_opened = False

        'If 

        'For Testing:
        If Raising_Orders.test_completed = True Then

            MsgBox("Testing completed - trade roundtrip completed")

        End If

        'For testing Git 2 tst
        'i_Raising_Orders.position_opened
        If Properties_Class.position_opened = True Then

            If ma_ticks_arr(0) <> Nothing And MID_price <> 0 Then

                ma_5_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6)) / 5
                ma_5_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5)) / 5
                'ma_3_ticks = (ma_3_ticks_arr(2) + ma_3_ticks_arr(1) + ma_3_ticks_arr(0)) / 3
                ma_10_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5) + ma_ticks_arr(4) + ma_ticks_arr(3) + ma_ticks_arr(2) + ma_ticks_arr(1)) / 10
                ma_10_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5) + ma_ticks_arr(4) + ma_ticks_arr(3) + ma_ticks_arr(2) + ma_ticks_arr(1) + ma_ticks_arr(0)) / 10

            End If

            i_Open_positions_cond_analysis_ma_5_10.ma_5_ticks_for_exit_calc = ma_5_ticks
            i_Open_positions_cond_analysis_ma_5_10.ma_10_ticks_for_exit_calc = ma_10_ticks
            i_Open_positions_cond_analysis_ma_5_10.price_for_exit_calc = ma_ticks_arr(10)
            thrd_open_pos_cond_analysis = New Thread(AddressOf i_Open_positions_cond_analysis_ma_5_10.open_positions_cond_analysis_ma_5_over_10)
            'For testing
            'TODO - To be removed for continuous trading
            If first_exit_trade = True Then

            Else
                thrd_open_pos_cond_analysis.Start(ma_ticks_arr)
                thrd_open_pos_cond_analysis.Join()
            End If

        End If

        'If ma_ticks_arr(7) <> Nothing And LAST_price <> "" Then

        'i_Raising_Orders.position_opened
        If Properties_Class.position_opened = False Then

            If ma_ticks_arr(0) <> Nothing And MID_price <> 0 Then

                'ma_5_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9)) / 2
                'ma_5_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8)) / 2
                'ma_10_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8)) / 3
                'ma_10_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7)) / 3

                ma_5_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6)) / 5
                ma_5_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5)) / 5
                'ma_3_ticks = (ma_3_ticks_arr(2) + ma_3_ticks_arr(1) + ma_3_ticks_arr(0)) / 3
                ma_10_ticks = (ma_ticks_arr(10) + ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5) + ma_ticks_arr(4) + ma_ticks_arr(3) + ma_ticks_arr(2) + ma_ticks_arr(1)) / 10
                ma_10_ticks_prior_per = (ma_ticks_arr(9) + ma_ticks_arr(8) + ma_ticks_arr(7) + ma_ticks_arr(6) + ma_ticks_arr(5) + ma_ticks_arr(4) + ma_ticks_arr(3) + ma_ticks_arr(2) + ma_ticks_arr(1) + ma_ticks_arr(0)) / 10


                '------Ensures that tick is a very first tick, when it is
                If end_very_first_time = False Then
                    very_first_time = True
                End If

                '-----Ensures that very first met condition is ignored, as position should only be entered after condition is met coming from a non-met state
                If ma_5_ticks < ma_10_ticks And very_first_time = True Then
                    ma_mixed_for_long_entry = True
                    end_very_first_time = True
                    very_first_time = False
                Else
                    If ma_5_ticks > ma_10_ticks Then
                        ma_mixed_for_short_entry = True
                    End If
                End If

                '-----Call conditions analysis calculations for long entry
                If ma_mixed_for_long_entry = True Then
                    'For testing
                    If first_trade = True Then

                    Else
                        Call i_Conditions_analysis_ma_5_10.conditions_analysis_ma_5_over_10(ma_5_ticks, ma_10_ticks, ma_5_ticks_prior_per, ma_10_ticks_prior_per, MID_price)
                    End If

                End If

                '-----Call conditions analysis calculations for short entry
                If ma_mixed_for_short_entry = True Then

                    'For testing
                    'Call i_Conditions_analysis_ma_5_10.conditions_analysis_ma_5_under_10(ma_5_ticks, ma_10_ticks, ma_5_ticks_prior_per, ma_10_ticks_prior_per, MID_price)

                End If

            End If

        End If

        'Erase VOLUME_TIME_SALES_arr

    End Sub


    Sub mm_size_return_handler(tickerId As Integer, field As Integer, size As Double)

        'Dim tick_type As String
        'Dim dt As DateTime
        'Dim VOLUME_TIME_SALES_arr() As String
        'Dim US_tm_6_hrs As TimeSpan
        'Dim ma_ticks_arr() As Double
        'Dim ma_3_ticks_arr() As Double
        'Dim ma_5_ticks_prior_per As Double
        'Dim ma_10_ticks_prior_per As Double
        'Dim thrd_open_pos_cond_analysis As Thread

        'ReDim ma_5_ticks_arr(5)
        'ReDim ma_3_ticks_arr(3)
        'ReDim Preserve ma_ticks_arr(10)

        Select Case field
            Case 1

            Case 2

            Case 6

            Case 7

            Case 9

        End Select

    End Sub

End Class
