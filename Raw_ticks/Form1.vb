Imports System.IO
Imports IBApi

Public Class Form1

    Public cnnt_cls As Connect = New Connect
    Dim current_time_cls As Current_time = New Current_time
    Dim data_cls As Data = New Data
    Dim orders_cls As Requesting_Open_Orders = New Requesting_Open_Orders
    Dim i_raising_orders As Raising_Orders = New Raising_Orders

    'Public WithEvents Wrapper_events As New EWrapperImpl
    Public Shared auto_trading As Boolean
    Public Shared str_wrt_opened As Boolean

    Private Sub btn_cnct_Click(sender As Object, e As EventArgs) Handles btn_cnct.Click

        Call cnnt_cls.connect()

    End Sub

    Private Sub btn_req_time_Click(sender As Object, e As EventArgs) Handles btn_req_time.Click

        Call current_time_cls.req_current_time()

    End Sub

    Private Sub btn_req_mkt_prc_Click(sender As Object, e As EventArgs) Handles btn_req_mkt_prc.Click

        auto_trading = False

        Call data_cls.req_data()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btn_req_TWS_opn_ords_Click(sender As Object, e As EventArgs) Handles btn_req_TWS_opn_ords.Click

        Call orders_cls.req_TWS_open_ords()

    End Sub

    Private Sub btn_req_opn_ords_Click(sender As Object, e As EventArgs) Handles btn_req_opn_ords.Click

        Call orders_cls.req_open_ords()

    End Sub

    Private Sub btn_order_Click(sender As Object, e As EventArgs) Handles btn_order.Click

        Dim action As String
        Dim quantity As Double
        Dim ord_prc As Double

        action = cbx_action.SelectedItem
        quantity = tbx_quantity.Text
        ord_prc = txb_price.Text

        Dim i_ordr As IBApi.Order
        i_ordr = New IBApi.Order

        i_ordr.Action = cbx_action.SelectedItem
        i_ordr.OrderType = "LMT"
        i_ordr.LmtPrice = txb_price.Text
        i_ordr.TotalQuantity = tbx_quantity.Text

        Call i_raising_orders.submit_ords(i_ordr)

    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click

        Dim orderID As Integer

        orderID = tbx_Order_ID.Text

        Call orders_cls.cancel_ords(orderID)

    End Sub

    Private Sub btn_req_ord_ID_Click(sender As Object, e As EventArgs) Handles btn_req_ord_ID.Click

        Connect.bool = True

        Call orders_cls.req_next_ord_ID()

    End Sub

    Private Sub btn_trade_last_tick_Click(sender As Object, e As EventArgs) Handles btn_trade_last_tick.Click

        auto_trading = True

        If Form1.str_wrt_opened = False Then
            Dim i_Open_str_wrt As Open_str_wrt = New Open_str_wrt
            Call i_Open_str_wrt.Open_str_wrt_method()
            Form1.str_wrt_opened = True
        End If

        Call data_cls.req_data()

    End Sub
End Class
