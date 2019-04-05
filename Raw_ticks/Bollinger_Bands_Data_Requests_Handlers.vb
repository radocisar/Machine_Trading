﻿Imports System.IO
Imports System.Threading

Public Class Bollinger_Bands_Data_Requests_Handlers

    Private BID_price As Double
    Private ASK_price As Double
    Private candle_arr() As Double
    Private Current_time As DateTime

    Sub mm_price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer)

        Dim upper_band As Double
        Dim lower_band As Double
        Dim middle_band As Double
        Dim open_price As Double
        Dim high_price As Double
        Dim low_price As Double
        Dim last_price As Double

        ReDim Preserve candle_arr(10)

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
                Exit Sub
        End Select



        Current_time = DateTime.Now

        If Current_time.Second <> 0 Then



        ElseIf Current_time.Second = 0 And last_price <> "" Then

            'assign candle into the array

        End If

        mls_200 = New TimeSpan(0, 0, 0, 0, 200)



        candle_arr(0) = candle_arr(1)
        candle_arr(1) = candle_arr(2)
        candle_arr(2) = candle_arr(3)
        candle_arr(3) = candle_arr(4)
        candle_arr(4) = candle_arr(5)
        candle_arr(5) = candle_arr(6)
        candle_arr(6) = candle_arr(7)
        candle_arr(7) = candle_arr(8)
        candle_arr(8) = candle_arr(9)
        candle_arr(9) = candle_arr(10)
        candle_arr(10) = last_price

    End Sub

End Class