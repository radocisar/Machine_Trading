Imports System.IO
Imports IBApi

Public Class Execute

    Dim i_raising_orders As Raising_Orders = New Raising_Orders
    'Dim i_Log_to_file As Log_to_file = New Log_to_file
    Public Shared position_opened_or_closed As String
    'Public Shared Event raise_ord(ByVal i_ordr As IBApi.Order)


    Sub execute_in_ord_LMT(ByVal i_Auto_open_trade_parameters As Auto_open_trade_parameters)

        'Dim action As String
        'Dim quantity As Double
        'Dim ord_prc As Double

        'action = i_Auto_open_trade_parameters.Auto_open_trade_action
        'quantity = i_Auto_open_trade_parameters.Auto_open_trade_size
        'ord_prc = i_Auto_open_trade_parameters.Auto_open_trade_price + 0.05

        Dim i_ordr As IBApi.Order
        i_ordr = New IBApi.Order

        i_ordr.Action = i_Auto_open_trade_parameters.Auto_open_trade_action
        i_ordr.OrderType = "LMT"

        If i_ordr.Action = "BUY" Then
            'For testing
            i_ordr.LmtPrice = i_Auto_open_trade_parameters.Auto_open_trade_price - 1 '+ 0.0005
        Else
            'For testing
            i_ordr.LmtPrice = i_Auto_open_trade_parameters.Auto_open_trade_price + 1 '- 0.0005
        End If

        i_ordr.TotalQuantity = i_Auto_open_trade_parameters.Auto_open_trade_size

        position_opened_or_closed = "Opened"

        ' This will only work for one opened position at the time
        'If i_raising_orders.position_opened <> True Then
        '    position_opened_or_closed = "Opened"
        'Else
        '    position_opened_or_closed = "Closed"
        'End If

        'Call i_raising_orders.submit_ords(action, quantity, ord_prc)

        '-----Logging - Auto Trading
        'Dim str_wrt As StreamWriter = New StreamWriter("C:\Raw_Data\Logging _Auto_Trading\log.txt", True)

        If i_ordr.Action = "BUY" Then
            'TODO Create a class/method for writting to file
            Log_to_file.main_lbl = "Opening Long Execution Details"
            Log_to_file.i_ordr_Action_lbl = "Action="
            Log_to_file.price_lbl = "Price="
            Log_to_file.i_ordr_Quantity_lbl = "Size="
            Log_to_file.i_ordr_OrderType_lbl = "Order Type="
            Log_to_file.Log_to_file_method(i_ordr.Action, i_ordr.LmtPrice, i_ordr.TotalQuantity, i_ordr.OrderType, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            'str_wrt.WriteLine("Opening Long Execution Details" & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
            '& i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)
            'str_wrt.WriteLine("Opening Long Execution Details" & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
            '& i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)

            Debug.Print("Establish Long - Execution Details:" & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
                                & i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)
        Else
            Log_to_file.main_lbl = "Opening Short Execution Details"
            Log_to_file.i_ordr_Action_lbl = "Action="
            Log_to_file.price_lbl = "Price="
            Log_to_file.i_ordr_Quantity_lbl = "Size="
            Log_to_file.i_ordr_OrderType_lbl = "Order Type="
            Log_to_file.Log_to_file_method(i_ordr.Action, i_ordr.LmtPrice, i_ordr.TotalQuantity, i_ordr.OrderType, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            'str_wrt.WriteLine("Opening Short Execution Details" & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
            '& i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)
            Debug.Print("Establish Short - Execution Details:" & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
                                & i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)

        End If

        'str_wrt.Close()

        Call i_raising_orders.submit_ords(i_ordr)

    End Sub


    Sub execute_out_ord_LMT(ByVal i_Auto_close_trade_parameters As Auto_close_trade_parameters)

        'Dim action As String
        'Dim quantity As Double
        'Dim ord_prc As Double

        'action = i_Auto_close_trade_parameters.Auto_close_trade_action
        'quantity = i_Auto_close_trade_parameters.Auto_close_trade_size
        'ord_prc = i_Auto_close_trade_parameters.Auto_close_trade_price - 0.05

        Dim i_ordr As IBApi.Order
        i_ordr = New IBApi.Order

        i_ordr.Action = i_Auto_close_trade_parameters.Auto_close_trade_action
        i_ordr.OrderType = "LMT"

        If i_ordr.Action = "SELL" Then
            'For testing
            i_ordr.LmtPrice = i_Auto_close_trade_parameters.Auto_close_trade_price + 0.5 '- 0.05
        Else
            'For testing
            i_ordr.LmtPrice = i_Auto_close_trade_parameters.Auto_close_trade_price - 0.5 '+ 0.05
        End If

        i_ordr.TotalQuantity = i_Auto_close_trade_parameters.Auto_close_trade_size

        position_opened_or_closed = "Closed"

        ' This will only work for one opened position at the time
        'If i_raising_orders.position_opened <> True Then
        '    position_opened_or_closed = "Opened"
        'Else
        '    position_opened_or_closed = "Closed"
        'End If

        'Call i_raising_orders.submit_ords(action, quantity, ord_prc)

        '-----Logging - Auto Trading
        'Dim str_wrt As StreamWriter = New StreamWriter("C:\Raw_Data\Logging _Auto_Trading\log.txt", True)

        If i_ordr.Action = "BUY" Then
            Log_to_file.main_lbl = "Short Exit Execution Details"
            Log_to_file.i_ordr_Action_lbl = "Action="
            Log_to_file.price_lbl = "Price="
            Log_to_file.i_ordr_Quantity_lbl = "Size="
            Log_to_file.i_ordr_OrderType_lbl = "Order Type="
            Log_to_file.Log_to_file_method(i_ordr.Action, i_ordr.LmtPrice, i_ordr.TotalQuantity, i_ordr.OrderType, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            'str_wrt.WriteLine("Short Exit Execution Details" & vbNewLine & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
            '& i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)
            Debug.Print("Exiting Short - Execution Details:" & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
                               & i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)

        Else
            Log_to_file.main_lbl = "Long Exit Execution Details"
            Log_to_file.i_ordr_Action_lbl = "Action="
            Log_to_file.price_lbl = "Price="
            Log_to_file.i_ordr_Quantity_lbl = "Size="
            Log_to_file.i_ordr_OrderType_lbl = "Order Type="
            Log_to_file.Log_to_file_method(i_ordr.Action, i_ordr.LmtPrice, i_ordr.TotalQuantity, i_ordr.OrderType, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            'str_wrt.WriteLine("Long Exit Execution Details" & vbNewLine & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
            '& i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)
            Debug.Print("Exiting Long - Execution Details:" & vbNewLine & "Contract=" & Functions.cntrt.Symbol & vbNewLine & "Action=" & i_ordr.Action & vbNewLine & "Price=" & i_ordr.LmtPrice & vbNewLine & "Size=" _
                               & i_ordr.TotalQuantity & vbNewLine & "Order Type=" & i_ordr.OrderType & vbNewLine)

        End If

        'str_wrt.Close()

        Call i_raising_orders.submit_ords(i_ordr)

    End Sub

End Class