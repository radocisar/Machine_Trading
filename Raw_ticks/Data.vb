﻿Imports IBApi
Imports System.IO

Public Class Data
    adsfgasdfasdfasdfasdfasdf
    'Dim wrapper_imp As EWrapperImpl = New EWrapperImpl
    Dim cls_Data_Requests_Handlers_write_to_file As Data_Requests_Handlers_write_to_file = New Data_Requests_Handlers_write_to_file
    Dim cls_Data_Requests_Handlers_ma_5_10 As Data_Requests_Handlers_ma_5_10 = New Data_Requests_Handlers_ma_5_10
    'Dim i_form1 As Form1 = New Form1
    'Public socket_client As EClientSocket = wrapper_imp.socketClient
    'Dim create_new_fldr_for_cntrt As Create_new_folder_for_cntrt = New Create_new_folder_for_cntrt

    Dim WithEvents wrapper_events As New EWrapperImpl
    Public file_path As String = "C:\Raw_Data\Raw_Tickstream\" & Functions.cntrt.Symbol

#Region "Data Requests Raisers"

    Public Sub req_data()


        If Form1.auto_trading = True Then

            'For testing
            Data_Requests_Handlers_ma_5_10.tst_data_time = Now

            Form1.cnnt_cls.socket_client.reqMktData(11, Functions.cntrt, "233", False, Nothing)


        Else
            'Form1.cnnt_cls.socket_client.reqMktData(11, Functions.cntrt, String.Empty, False, Nothing)
            Call Create_new_folder_for_cntrt.file_path()

            '----VOLUME_TIME_SALES type of generic tick:
            Form1.cnnt_cls.socket_client.reqMktData(11, Functions.cntrt, "233", False, Nothing)

            '----VOLUME_TIME_SALES type of generic tick (does not seem to work for stocks):
            'YForm1.cnnt_cls.socket_client.reqMktData(11, Functions.cntrt, "375", False, Nothing)
        End If

    End Sub

#End Region


#Region "Data Requests Handlers"


    Sub price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer) Handles wrapper_events.on_price_return

        If Form1.auto_trading = True Then

            '----- For FX data handling for trading
            Call cls_Data_Requests_Handlers_ma_5_10.mm_price_return_handler(tickerId, field, price, canAutoExecute)

        Else

            '----- For all instruments' data handling for raw ticks download
            Call cls_Data_Requests_Handlers_write_to_file.m_price_return_handler(tickerId, field, price, canAutoExecute)
            'Call cls_Data_Requests_Handlers_ma_5_10.mm_price_return_handler(tickerId As Integer, field As Integer, price As Double, canAutoExecute As Integer)

        End If

    End Sub

    Sub size_return_handler(tickerId As Integer, field As Integer, size As Double) Handles wrapper_events.on_size_return

        If Form1.auto_trading = True Then

            '----- For FX data handling for trading
            Call cls_Data_Requests_Handlers_ma_5_10.mm_size_return_handler(tickerId, field, size)

        Else

            '----- For all instruments' data handling for raw ticks download
            Call cls_Data_Requests_Handlers_write_to_file.m_size_return_handler(tickerId, field, size)
            'Call cls_Data_Requests_Handlers_ma_5_10.mm_size_return_handler(tickerId, field, size)

        End If

    End Sub

    Sub string_return_handler(tickerId As Integer, field As Integer, string_return As String) Handles wrapper_events.on_string_return

        If Form1.auto_trading = True Then

            '----- For Stocks data handling for trading
            Call cls_Data_Requests_Handlers_ma_5_10.mm_string_return_handler(tickerId, field, string_return)

        Else

            '----- For all instruments' data handling for raw ticks download
            Call cls_Data_Requests_Handlers_write_to_file.m_string_return_handler(tickerId, field, string_return)

        End If

    End Sub

#End Region


End Class
