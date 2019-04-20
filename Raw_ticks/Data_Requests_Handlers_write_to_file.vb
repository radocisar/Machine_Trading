Imports System.IO

Public Class Data_Requests_Handlers_write_to_file

    'Dim dta As Data = New Data
    Public bln As Boolean
    Public file_txt As String = "Raw_ticks.txt"

    Sub m_price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer)

        Dim sw As StreamWriter
        Dim tick_type As String

        'prc = AxTWSLib._DTwsEvents_tickPriceEvent
        'MsgBox(eventArgs.price & vbCrLf & eventArgs.tickType & vbCrLf & eventArgs.id & vbCrLf & eventArgs.canAutoExecute)

        'prc = price

        If File.Exists(Create_new_folder_for_cntrt.file_path & file_txt) = False Then
            bln = True
        End If

        Select Case field
            Case 1
                'tick_type = "BID"
                tick_type = "B"
            Case 2
                'tick_type = "ASK"
                tick_type = "A"
            Case 4
                'tick_type = "LAST"
                tick_type = "L"
            Case 6
                'tick_type = "HIGH"
                tick_type = "H"
            Case 7
                'tick_type = "LOW"
                tick_type = "L"
            Case 9
                'tick_type = "CLOSE"
                tick_type = "C"
            Case 14
                'tick_type = "OPEN"
                tick_type = "O"
            Case 15
                'tick_type = "13_WK_LOW"
                tick_type = "13_W_L"
            Case 16
                'tick_type = "13_WK_HIGH"
                tick_type = "13_W_H"
            Case 17
                'tick_type = "26_WK_LOW"
                tick_type = "26_W_L"
            Case 18
                'tick_type = "26_WK_HIGH"
                tick_type = "26_W_H"
            Case 19
                'tick_type = "52_WK_LOW"
                tick_type = "52_W_L"
            Case 20
                'tick_type = "52_WK_HIGH"
                tick_type = "52_W_H"
            Case 35
                'tick_type = "AUCTION_PRICE"
                tick_type = "AU_P"
            Case 37
                'tick_type = "MARK_PRICE"
                tick_type = "M_P"
            Case 50
                'tick_type = "BID_YLD"
                tick_type = "B_Y"
            Case 51
                'tick_type = "ASK_YLD"
                tick_type = "A_Y"
            Case 52
                'tick_type = "LAST_YLD"
                tick_type = "L_Y"
            Case 57
                'tick_type = "LAST_RTH"
                tick_type = "L_R"
            Case 66
                'tick_type = "DELAYED_BID"
                tick_type = "D_B"
            Case 67
                'tick_type = "DELAYED_ASK"
                tick_type = "D_A"
            Case 68
                'tick_type = "DELAYED_LAST"
                tick_type = "D_L"
            Case 72
                'tick_type = "DELAYED_HIGH"
                tick_type = "D_H"
            Case 73
                'tick_type = "DELAYED_LOW"
                tick_type = "D_L"
            Case 75
                'tick_type = "DELAYED_CLOSE"
                tick_type = "D_C"
            Case 76
                'tick_type = "DELAYED_OPEN"
                tick_type = "D_O"
            Case Else
                tick_type = "Unknown"
        End Select

        'cntr_id = tickerId
        'can_auto_exec = canAutoExecute

        sw = New StreamWriter(Create_new_folder_for_cntrt.file_path & file_txt, True)

        sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & price)

        sw.Close()

    End Sub


    Sub m_size_return_handler(tickerId As Integer, field As Integer, size As Double)

        Dim sw As StreamWriter
        Dim tick_type As String

        'prc = AxTWSLib._DTwsEvents_tickPriceEvent
        'MsgBox(eventArgs.price & vbCrLf & eventArgs.tickType & vbCrLf & eventArgs.id & vbCrLf & eventArgs.canAutoExecute)

        'prc = price

        If File.Exists(Create_new_folder_for_cntrt.file_path & file_txt) = False Then
            bln = True
        End If

        Select Case field
            Case 0
                'tick_type = "BID_SIZE"
                tick_type = "B_S"
            Case 3
                'tick_type = "ASK_SIZE"
                tick_type = "A_S"
            Case 5
                'tick_type = "LAST_SIZE"
                tick_type = "L_S"
            Case 8
                'tick_type = "VOLUME"
                tick_type = "V"
            Case 21
                'tick_type = "AVG_VOLUME"
                tick_type = "A_V"
            Case 22
                'tick_type = "OPEN_INT"
                tick_type = "O_I"
            Case 27
                'tick_type = "CALL_OPEN_INT"
                tick_type = "C_O_I"
            Case 28
                'tick_type = "PUT_OPEN_INT"
                tick_type = "P_O_I"
            Case 29
                'tick_type = "CALL_VOLUME"
                tick_type = "C_V"
            Case 30
                'tick_type = "PUT_VOLUME"
                tick_type = "P_V"
            Case 34
                'tick_type = "AUCTION_VOLUME"
                tick_type = "AU_V"
            Case 36
                'tick_type = "AUCTION_IMBALANCE"
                tick_type = "AU_IM"
            Case 54
                'tick_type = "TRADE_COUNT_DAY"
                tick_type = "T_C_D"
            Case 55
                'tick_type = "TRADE_COUNT_MINUTE"
                tick_type = "T_C_M"
            Case 56
                'tick_type = "VOLUME_MINUTE"
                tick_type = "V_M"
            Case 61
                'tick_type = "REGULATORY_IMBALANCE"
                tick_type = "R_IM"
            Case 63
                'tick_type = "VOLUME_3_MINS"
                tick_type = "V_3_M"
            Case 64
                'tick_type = "VOLUME_5_MINS"
                tick_type = "V_5_M"
            Case 65
                'tick_type = "VOLUME_10_MINS"
                tick_type = "V_10_M"
            Case 69
                'tick_type = "DELAYED_BID_SIZE"
                tick_type = "D_B_S"
            Case 70
                'tick_type = "DELAYED_ASK_SIZE"
                tick_type = "D_A_S"
            Case 71
                'tick_type = "DELAYED_LAST_SIZE"
                tick_type = "D_L_S"
            Case 74
                'tick_type = "DELAYED_VOLUME"
                tick_type = "D_V"
            Case Else
                tick_type = "Unknown"
        End Select

        'cntr_id = tickerId
        'can_auto_exec = canAutoExecute

        'sw = New StreamWriter("C:\Users\Rados\Documents\Visual Studio 2015\Projects\Raw_ticks\Raw_ticks\Raw_ticks.txt", True)
        sw = New StreamWriter(Create_new_folder_for_cntrt.file_path & file_txt, True)

        sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & size)

        sw.Close()

    End Sub


    Sub m_string_return_handler(tickerId As Integer, field As Integer, string_return As String)

        Dim sw As StreamWriter
        Dim tick_type As String
        Dim dt As DateTime
        Dim VOLUME_TIME_SALES_arr() As String
        Dim US_tm_6_hrs As TimeSpan

        'Dim long_dt As Int64

        'prc = AxTWSLib._DTwsEvents_tickPriceEvent
        'MsgBox(eventArgs.price & vbCrLf & eventArgs.tickType & vbCrLf & eventArgs.id & vbCrLf & eventArgs.canAutoExecute)

        'prc = price

        If File.Exists(Create_new_folder_for_cntrt.file_path & file_txt) = False Then
            bln = True
        End If

        sw = New StreamWriter(Create_new_folder_for_cntrt.file_path & file_txt, True)

        Select Case field
            Case 32
                'tick_type = "BID_EXCH"
                tick_type = "B_E"
                sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & string_return)
            Case 33
                'tick_type = "ASK_EXCH"
                tick_type = "A_E"

                sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & string_return)
            Case 45
                'tick_type = "LAST_TIMESTAMP"
                tick_type = "L_T"

                dt = New DateTime(1970, 1, 1, 0, 0, 0)
                dt = dt.AddSeconds(CInt(string_return))

                If bln = True Then
                    sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & dt)
                    sw.WriteLine("----------------------------------------------------")

                    bln = False

                Else
                    sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & dt)
                End If

            'Case 47
            '    tick_type = "FUND_RATIOS"
            '    sw.WriteLine(tick_type & ": " & string_return)
            Case 48
                'tick_type = "VOLUME_TIME_SALES"

                VOLUME_TIME_SALES_arr = string_return.Split(";")

                dt = New DateTime(1970, 1, 1, 0, 0, 0)

                dt = dt.AddMilliseconds(CLng(VOLUME_TIME_SALES_arr(2)))
                US_tm_6_hrs = New TimeSpan(5, 0, 0)

                sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & "T_S|" & (dt - US_tm_6_hrs).ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & VOLUME_TIME_SALES_arr(0) & "|" & VOLUME_TIME_SALES_arr(1) &
                             "|" & VOLUME_TIME_SALES_arr(4) & "|" & VOLUME_TIME_SALES_arr(3) * 100 & "|")

                'sw.WriteLine(tick_type & "- Price: " & VOLUME_TIME_SALES_arr(0))
                'sw.WriteLine(tick_type & "- Size: " & VOLUME_TIME_SALES_arr(1))
                'sw.WriteLine(tick_type & "- Time: " & dt)
                'sw.WriteLine(tick_type & "- Total_Volume: " & VOLUME_TIME_SALES_arr(3) * 100)
                'sw.WriteLine(tick_type & "- VWAP: " & VOLUME_TIME_SALES_arr(4))
                'sw.WriteLine(tick_type & "- Single Mmkr: " & VOLUME_TIME_SALES_arr(5))

                Erase VOLUME_TIME_SALES_arr

                'Case 59
                '    tick_type = "DIVDS"
                '    sw.WriteLine(tick_type & ": " & string_return)       
                'Case 62
                '    tick_type = "NEWS"
                '    sw.WriteLine(tick_type & ": " & string_return)
            Case 77
                'tick_type = "VOLUME_TRADE_TIME_SALES"
                tick_type = "T_T_S"
                sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & string_return)
            Case Else
                tick_type = "Unknown"
                sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & "|" & tick_type & "|" & string_return)
        End Select

        'cntr_id = tickerId
        'can_auto_exec = canAutoExecute

        sw.Close()

    End Sub


End Class
