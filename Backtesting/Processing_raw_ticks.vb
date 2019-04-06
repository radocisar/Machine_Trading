Imports System.IO

Public Class Processing_raw_ticks

    Public file_path As String = "E:\Test_FX_Ticks_Download\"
    Public file_txt As String = "Raw_ticks_1.txt"

    Public last As Double?
    Public bid As Double?
    Public ask As Double?
    Public tm As DateTime
    Public last_size As Integer?
    Public bid_size As Integer?
    Public ask_size As Integer?
    Public ttl_v As Integer?
    Public VWAP As Double?
    Public prev_close As Double?
    Public open As Double?
    Public day_low As Double?
    Public day_high As Double?

    Public last_coll As New List(Of Double?)(440000)
    Public bid_coll As New List(Of Double?)(440000)
    Public ask_coll As New List(Of Double?)(440000)
    Public tm_coll As New List(Of DateTime)(440000)
    Public last_size_coll As New List(Of Integer?)(440000)
    Public bid_size_coll As New List(Of Integer?)(440000)
    Public ask_size_coll As New List(Of Integer?)(440000)
    Public ttl_v_coll As New List(Of Integer?)(440000)
    Public VWAP_coll As New List(Of Double?)(440000)

    Public bln As Boolean
    'Public master_arr(8) As Double

    Public Sub moving_average_calc()

        'Call assigning_values_to_master_arr()

        ' Test the moving average back testing and then make this into a .dll processing raw data file into an array




        'If last <> 0 Then

        '    'If it is not a recurring trade in a moving average array
        '    If bln = True Then

        '        last_arr(0) = last_arr(1)
        '        last_arr(1) = last_arr(2)
        '        last_arr(2) = last_arr(3)
        '        last_arr(3) = last_arr(4)
        '        last_arr(4) = last

        '    Else

        '        If last_arr(0) = 0 Then
        '            last_arr(0) = last
        '        ElseIf last_arr(1) = 0 Then
        '            last_arr(1) = last
        '        ElseIf last_arr(2) = 0 Then
        '            last_arr(2) = last
        '        ElseIf last_arr(3) = 0 Then
        '            last_arr(3) = last
        '        Else
        '            last_arr(4) = last
        '            bln = True
        '        End If

        '    End If

        'End If

        'last = 0




        'bln = False

    End Sub


    Public Function calc_moving_averages(ByVal last As Double) As Double

        'Dim avg_3_per As Double
        'Dim avg_5_per As Double

        'Return 1

        ' = calc_moving_averages

    End Function

    Public Sub processing_raw_ticks_file()

        Dim sr As StreamReader
        Dim sw As StreamWriter
        Dim line_arr() As String
        Dim strt_of_trd_day As DateTime
        Dim end_of_trd_day As DateTime
        Dim mls_200 As TimeSpan
        'Dim mls_1 As TimeSpan
        Dim tm_count As DateTime
        Dim initial_empty_tms_filled As Boolean
        Dim intraday_empty_tms_filled As Boolean
        Dim final_empty_tms_filled As Boolean

        'last = Nothing
        'last_size = Nothing
        'ttl_v = Nothing
        'VWAP = Nothing
        'bid = Nothing
        'bid_size = Nothing
        'ask = Nothing
        'ask_size = Nothing


        sr = New StreamReader(file_path & file_txt)

        '-----Start of trading day for US stocks

        Do Until sr.ReadLine.Contains("-----")
            'Ensures the actual collection population only begins from after the initial paint
        Loop

        line_arr = Split(sr.ReadLine, "|")

        tm = Convert.ToDateTime(line_arr(0))

        strt_of_trd_day = tm.Month & "/" & tm.Day & "/" & tm.Year & " " & #00:00:00#

        mls_200 = New TimeSpan(0, 0, 0, 0, 200)
        tm_count = strt_of_trd_day + mls_200

        Do Until sr.EndOfStream = True

            'Convert.ToDateTime("01/01/1900 00:00:00.200")

            'Debug.Print(mls_200.ToString())
            'Debug.Print(mls_200.ToString("MM/dd/yyyy HH:mm:ss.ffff"))

            '-----Initial filling of empty times from start of day until first recorded tick
            Do Until initial_empty_tms_filled = True

                If mls_200.CompareTo(tm - (tm_count - mls_200)) < 0 Then

                    last_coll.Add(last)
                    tm_coll.Add(tm_count)
                    last_size_coll.Add(last_size)
                    ttl_v_coll.Add(ttl_v)
                    VWAP_coll.Add(VWAP)
                    bid_coll.Add(bid)
                    bid_size_coll.Add(bid_size)
                    ask_coll.Add(ask)
                    ask_size_coll.Add(ask_size)

                    tm_count = tm_count + mls_200

                Else

                    initial_empty_tms_filled = True

                End If

            Loop


            '------ Assign values to variable and then to the collections

            Call assigning_values_to_master_arr(intraday_empty_tms_filled, mls_200, tm_count, line_arr)


            intraday_empty_tms_filled = False

            Erase line_arr
            line_arr = Split(sr.ReadLine, "|")

            tm = Convert.ToDateTime(line_arr(0))


            If tm.CompareTo(tm_count) < 1 Then

                If line_arr(1) <> "L" And line_arr(1) <> "L_S" And line_arr(1) <> "V" Then
                    last_coll.RemoveAt(last_coll.Count - 1)
                    tm_coll.RemoveAt(tm_coll.Count - 1)
                    last_size_coll.RemoveAt(last_size_coll.Count - 1)
                    ttl_v_coll.RemoveAt(ttl_v_coll.Count - 1)
                    VWAP_coll.RemoveAt(VWAP_coll.Count - 1)
                    bid_coll.RemoveAt(bid_coll.Count - 1)
                    bid_size_coll.RemoveAt(bid_size_coll.Count - 1)
                    ask_coll.RemoveAt(ask_coll.Count - 1)
                    ask_size_coll.RemoveAt(ask_size_coll.Count - 1)
                Else

                End If

            Else

                tm_count = tm_count + mls_200

            End If

        Loop

        sr.Close()


        '-----Append last record in the raw ticks file to the processed ticks file (as the sr.EndOfStream precludes last line to be assigned to the collections as the loop quits before it gets a chance to assign the last record)

        Call assigning_values_to_master_arr(intraday_empty_tms_filled, mls_200, tm_count, line_arr)


        '-----Append all missing times until 16:00 EST

        end_of_trd_day = tm_count.Month & "/" & tm_count.Day & "/" & tm_count.Year & " " & #23:59:59#

        If tm_count <> end_of_trd_day Then

            tm_count = tm_count + mls_200

        End If

        'mls_1 = New TimeSpan(0, 0, 0, 0, 1)

        Do Until final_empty_tms_filled = True

            If tm_count <= end_of_trd_day And tm_coll(tm_coll.Count - 1) <> end_of_trd_day Then

                last_coll.Add(last)
                tm_coll.Add(tm_count)
                last_size_coll.Add(last_size)
                ttl_v_coll.Add(ttl_v)
                VWAP_coll.Add(VWAP)
                bid_coll.Add(bid)
                bid_size_coll.Add(bid_size)
                ask_coll.Add(ask)
                ask_size_coll.Add(ask_size)

                tm_count = tm_count + mls_200

            Else

                final_empty_tms_filled = True

            End If

        Loop


        '-----Write all the collections into a file

        sw = New StreamWriter(file_path & "Processed_Ticks.txt", True)

        For x = 0 To last_coll.Count - 1
            sw.WriteLine(tm_coll(x).ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & Format(bid_coll(x), "0.00") & "|" & bid_size_coll(x) * 100 & "|" & Format(ask_coll(x), "0.00") & "|" & ask_size_coll(x) * 100 & "|" & Format(last_coll(x), "0.00") _
            & "|" & last_size_coll(x) * 100 & "|" & Format(VWAP_coll(x), "0.00000000") & "|" & ttl_v_coll(x))
        Next x

        sw.Close()

    End Sub

    Private Sub assigning_values_to_master_arr(ByRef intraday_empty_tms_filled As Boolean, ByRef mls_200 As TimeSpan, ByRef tm_count As DateTime, ByRef line_arr() As String)

        Do Until intraday_empty_tms_filled = True

            If mls_200.CompareTo(tm - (tm_count - mls_200)) < 0 Then

                last_coll.Add(last)
                tm_coll.Add(tm_count)
                last_size_coll.Add(last_size)
                ttl_v_coll.Add(ttl_v)
                VWAP_coll.Add(VWAP)
                bid_coll.Add(bid)
                bid_size_coll.Add(bid_size)
                ask_coll.Add(ask)
                ask_size_coll.Add(ask_size)

                tm_count = tm_count + mls_200

            Else

                intraday_empty_tms_filled = True

            End If

        Loop


        Select Case line_arr(1)

            Case "T_S"
                '-----populate fields with available values
                Try
                    last = CDbl(line_arr(3))

                    'tm = Convert.ToDateTime(line_arr(2))
                    last_size = CInt(line_arr(4))
                    ttl_v = CInt(line_arr(6))
                    VWAP = CDbl(line_arr(5))

                    '-----populate collections with available data
                    last_coll.Add(last)
                    tm_coll.Add(tm_count)
                    last_size_coll.Add(last_size)
                    ttl_v_coll.Add(ttl_v)
                    VWAP_coll.Add(VWAP)

                    '-----populate rest of collections with copied data from previous times
                    bid_coll.Add(bid)
                    bid_size_coll.Add(bid_size)
                    ask_coll.Add(ask)
                    ask_size_coll.Add(ask_size)

                Catch ex As System.InvalidCastException

                    If line_arr(3) = "" And CDbl(line_arr(5)) = VWAP_coll(VWAP_coll.Count - 1) And CDbl(line_arr(6)) = ttl_v_coll(ttl_v_coll.Count - 1) Then

                        '-----populate rest of the collections with copied data from previous times
                        last_coll.Add(last)
                        tm_coll.Add(tm_count)
                        last_size_coll.Add(last_size)
                        ttl_v_coll.Add(ttl_v)
                        VWAP_coll.Add(VWAP)
                        bid_coll.Add(bid)
                        bid_size_coll.Add(bid_size)
                        ask_coll.Add(ask)
                        ask_size_coll.Add(ask_size)

                        Exit Try
                    End If

                End Try

            Case "B"

                Try
                    '-----populate fields with available values
                    bid = CDbl(line_arr(2))
                    'tm = Convert.ToDateTime(line_arr(0))

                    '-----populate collections with available data
                    bid_coll.Add(bid)

                    '-----populate rest of collections with copied data from previous times
                    bid_size_coll.Add(bid_size)
                    ask_coll.Add(ask)
                    ask_size_coll.Add(ask_size)
                    last_coll.Add(last)
                    tm_coll.Add(tm_count)
                    last_size_coll.Add(last_size)
                    ttl_v_coll.Add(ttl_v)
                    VWAP_coll.Add(VWAP)

                Catch ex As System.InvalidCastException

                    If line_arr(2) = "" Then

                        Exit Try

                    End If

                End Try

            Case "A"

                Try

                    '-----populate fields with available values
                    ask = CDbl(line_arr(2))
                    'tm = Convert.ToDateTime(line_arr(0))

                    '-----populate collections with available data
                    ask_coll.Add(ask)

                    '-----populate rest of collections with copied data from previous times
                    bid_coll.Add(bid)
                    bid_size_coll.Add(bid_size)
                    ask_size_coll.Add(ask_size)
                    last_coll.Add(last)
                    tm_coll.Add(tm_count)
                    last_size_coll.Add(last_size)
                    ttl_v_coll.Add(ttl_v)
                    VWAP_coll.Add(VWAP)

                Catch ex As System.InvalidCastException

                    If line_arr(2) = "" Then

                        Exit Try

                    End If

                End Try

            Case "B_S"

                Try

                    '-----populate fields with available values
                    bid_size = CInt(line_arr(2))
                    'tm = Convert.ToDateTime(line_arr(0))

                    '-----populate collections with available data
                    bid_size_coll.Add(bid_size)

                    '-----populate rest of collections with copied data from previous times
                    bid_coll.Add(bid)
                    ask_coll.Add(ask)
                    ask_size_coll.Add(ask_size)
                    last_coll.Add(last)
                    tm_coll.Add(tm_count)
                    last_size_coll.Add(last_size)
                    ttl_v_coll.Add(ttl_v)
                    VWAP_coll.Add(VWAP)

                Catch ex As System.InvalidCastException

                    If line_arr(2) = "" Then

                        Exit Try

                    End If

                End Try

            Case "A_S"

                Try

                    '-----populate fields with available values
                    ask_size = CInt(line_arr(2))
                    'tm = Convert.ToDateTime(line_arr(0))

                    '-----populate collections with available data
                    ask_size_coll.Add(ask_size)

                    '-----populate rest of collections with copied data from previous times
                    bid_coll.Add(bid)
                    bid_size_coll.Add(bid_size)
                    ask_coll.Add(ask)
                    last_coll.Add(last)
                    tm_coll.Add(tm_count)
                    last_size_coll.Add(last_size)
                    ttl_v_coll.Add(ttl_v)
                    VWAP_coll.Add(VWAP)

                Catch ex As System.InvalidCastException

                    If line_arr(2) = "" Then

                        Exit Try

                    End If

                End Try

                'Case "L"
                '    last = cdbl(line_arr(2))
                '    tm = Convert.ToDateTime(line_arr(0))

                'Case "L_S"
                '    '-----populate fields with available values
                '    last_size = CInt(line_arr(2))
                '    'tm = Convert.ToDateTime(line_arr(0))

                '    '-----populate collections with available data
                '    last_size_coll.Add(last_size)

                '    '-----populate rest of collections with copied data from previous times
                '    bid_coll.Add(bid)
                '    bid_size_coll.Add(bid_size)
                '    ask_coll.Add(ask)
                '    ask_size_coll.Add(ask_size)
                '    last_coll.Add(last)
                '    tm_coll.Add(tm_count)
                '    ttl_v_coll.Add(ttl_v)
                '    VWAP_coll.Add(VWAP)

                'Case "V"
                '    '-----populate fields with available values
                '    ttl_v = CInt(line_arr(2))
                '    'tm = Convert.ToDateTime(line_arr(0))

                '    '-----populate collections with available data
                '    ttl_v_coll.Add(ttl_v)

                '    '-----populate rest of collections with copied data from previous times
                '    bid_coll.Add(bid)
                '    bid_size_coll.Add(bid_size)
                '    ask_coll.Add(ask)
                '    ask_size_coll.Add(ask_size)
                '    last_coll.Add(last)
                '    tm_coll.Add(tm_count)
                '    last_size_coll.Add(last_size)
                '    VWAP_coll.Add(VWAP)

        End Select

    End Sub

End Class
