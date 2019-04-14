Imports System.IO
Imports IBApi

Public Class Raising_Orders

    Private WithEvents Wrapper_events As New EWrapperImpl
    Private WithEvents Exec_wrapper_events As New Exec_EWrapperImpl
    'Public position_opened As Boolean
    Dim i_Requesting_Open_Orders As Requesting_Open_Orders = New Requesting_Open_Orders
    'Dim i_Data_Requests_Handlers_ma_5_10 As Data_Requests_Handlers_ma_5_10 = New Data_Requests_Handlers_ma_5_10
    'Dim i_Properties_Class As Properties_Class = New Properties_Class
    Dim i_connect As Connect = New Connect
    Dim i_EWrapperImpl As EWrapperImpl = New EWrapperImpl
    Public Shared Order_ID As Integer

    Public Shared first_order_submited As Boolean
    'For Testing
    Public Shared test_completed As Boolean

    Public Event on_current_time_1(ByVal time As Long)

    '-----Logging - Auto Trading
    'Dim str_wrt_4 As StreamWriter = New StreamWriter("C:\Raw_Data\Logging _Auto_Trading\log.txt", True)

#Region "Orders Raisers"

    Public Sub submit_ords(ByVal i_ordr As IBApi.Order)

        Order_ID = Connect.wrapper_imp.nextOrderId

        'TODO Implement proper NextOrderID requesting for next order
        'NOTE Be carefull, as when a session (run) is interupted the first_order_submited will reset to false, even if an order had been submitted during prior session (run)
        If first_order_submited = True Then
            Order_ID = Order_ID + 1
        End If

        'TODO Implement NextOrderID retrieval and assignment for closing order
        first_order_submited = True
        Debug.Print("Submit Ordr Order ID: " & Order_ID)
        'For testing
        'Order_ID = 81
        'Order_ID = Order_ID + 1



        Form1.cnnt_cls.socket_client.placeOrder(Order_ID, Functions.cntrt, i_ordr)

    End Sub

    Public Sub cancel_ords(ByVal orderID As Integer)

        Form1.cnnt_cls.socket_client.cancelOrder(orderID)

    End Sub
#End Region

#Region "Orders Handlers"

    Sub ords_stts_handler(orderId As Integer, status As String, filled As Double, remaining As Double, avgFillPrice As Double, permId As Integer,
                          parentId As Integer, lastFillPrice As Double, clientId As Integer, whyHeld As String) Handles Wrapper_events.on_ords_status

        'MsgBox("Order Status:" & vbCrLf & "OrderStatus Id: " & orderId & vbCrLf & "Status: " & status & vbCrLf & "Filled: " & filled & vbCrLf & "Remaining: " & remaining &
        'vbCrLf & "AvgFillPrice: " & avgFillPrice & vbCrLf & "PermId: " & permId & vbCrLf & "ParentId: " & parentId & vbCrLf & "LastFillPrice: " & lastFillPrice &
        'vbCrLf & "ClientId: " & clientId & vbCrLf & "WhyHeld: " & whyHeld)

        '-----Logging - Auto Trading        
        'Dim str_wrt As StreamWriter = New StreamWriter("C:\Raw_Data\Logging _Auto_Trading\log.txt", True)

        If Form1.str_wrt_opened = False Then
            Dim i_Open_str_wrt As Open_str_wrt = New Open_str_wrt
            Call i_Open_str_wrt.Open_str_wrt_method()
            Form1.str_wrt_opened = True
        End If

        'str_wrt.WriteLine("Order ID=" & orderId & vbNewLine & "Status=" & status & vbNewLine)
        Log_to_file.main_lbl = "Order Execution Details"
        Log_to_file.orderId_lbl = "Order ID="
        Log_to_file.status_lbl = "Status="
        Log_to_file.Log_to_file_method(Nothing, Nothing, Nothing, Nothing, orderId, status, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)

        Debug.Print("Order ID=" & orderId & vbNewLine & "Status=" & status & vbNewLine)

        'str_wrt.Close()

        'TODO Add filled order confirmation "status" And ensure this Is only considered upon opening of a New position
        If status = "Submitted" Or status = "PreSubmitted" And Execute.position_opened_or_closed = "Opened" Then

            'position_opened = True
            Properties_Class.position_opened = True
        End If
        'TODO Add filled order confirmation "status" and ensure this is only considered upon exiting of position
        If status = "submitted" Or status = "PreSubmitted" And Execute.position_opened_or_closed = "Closed" Then
            'For testing
            'Properties_Class.position_opened = False
            test_completed = True
        End If

    End Sub

    Sub exec_details_handler(reqId As Integer, contract As IBApi.Contract, execution As IBApi.Execution) Handles Wrapper_events.on_exec_details

        MsgBox("ExecDetails:" & vbCrLf & "Req ID: " & reqId & vbCrLf & contract.Symbol & vbCrLf & contract.SecType & vbCrLf & "Price: " & execution.Price & vbCrLf & execution.Exchange)

    End Sub

    Sub cmssn_rpt_handler(commissionReport As IBApi.CommissionReport) Handles Wrapper_events.on_cmssn_rpt

        MsgBox("CommissionReport:" & vbCrLf & "Commission: " & commissionReport.Currency & " " & commissionReport.Commission)

    End Sub

#End Region

End Class
