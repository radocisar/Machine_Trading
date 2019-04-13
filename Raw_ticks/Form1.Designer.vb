<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btn_cnct = New System.Windows.Forms.Button()
        Me.btn_req_time = New System.Windows.Forms.Button()
        Me.btn_req_mkt_prc = New System.Windows.Forms.Button()
        Me.lbl_connection = New System.Windows.Forms.Label()
        Me.btn_req_TWS_opn_ords = New System.Windows.Forms.Button()
        Me.btn_req_opn_ords = New System.Windows.Forms.Button()
        Me.tbx_quantity = New System.Windows.Forms.TextBox()
        Me.txb_price = New System.Windows.Forms.TextBox()
        Me.cbx_action = New System.Windows.Forms.ComboBox()
        Me.btn_order = New System.Windows.Forms.Button()
        Me.lbl_action = New System.Windows.Forms.Label()
        Me.lbl_quantity = New System.Windows.Forms.Label()
        Me.lbl_price = New System.Windows.Forms.Label()
        Me.btn_cancel = New System.Windows.Forms.Button()
        Me.tbx_Order_ID = New System.Windows.Forms.TextBox()
        Me.lbl_Order_ID = New System.Windows.Forms.Label()
        Me.btn_req_ord_ID = New System.Windows.Forms.Button()
        Me.btn_trade_last_tick = New System.Windows.Forms.Button()
        Me.lbl_requesting_open_orders = New System.Windows.Forms.Label()
        Me.lbl_downloading_tick_data = New System.Windows.Forms.Label()
        Me.lbl_auto_trading = New System.Windows.Forms.Label()
        Me.lbl_candle_resolution = New System.Windows.Forms.Label()
        Me.tbx_candle_resolution = New System.Windows.Forms.TextBox()
        Me.lbl_for_bollinger_bands = New System.Windows.Forms.Label()
        Me.tbx_std_dev = New System.Windows.Forms.TextBox()
        Me.lbl_std_dev = New System.Windows.Forms.Label()
        Me.lbl_min_num_of_candles = New System.Windows.Forms.Label()
        Me.tbx_min_num_of_candles = New System.Windows.Forms.TextBox()
        Me.lbl_lookbck_for_calcs = New System.Windows.Forms.Label()
        Me.tbx_lookback_for_calcs = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btn_cnct
        '
        Me.btn_cnct.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.btn_cnct.Location = New System.Drawing.Point(6, 43)
        Me.btn_cnct.Margin = New System.Windows.Forms.Padding(2)
        Me.btn_cnct.Name = "btn_cnct"
        Me.btn_cnct.Size = New System.Drawing.Size(63, 21)
        Me.btn_cnct.TabIndex = 0
        Me.btn_cnct.Text = "Connect"
        Me.btn_cnct.UseVisualStyleBackColor = True
        '
        'btn_req_time
        '
        Me.btn_req_time.Location = New System.Drawing.Point(73, 43)
        Me.btn_req_time.Margin = New System.Windows.Forms.Padding(2)
        Me.btn_req_time.Name = "btn_req_time"
        Me.btn_req_time.Size = New System.Drawing.Size(107, 26)
        Me.btn_req_time.TabIndex = 1
        Me.btn_req_time.Text = "Req_Current_Time"
        Me.btn_req_time.UseVisualStyleBackColor = True
        '
        'btn_req_mkt_prc
        '
        Me.btn_req_mkt_prc.Location = New System.Drawing.Point(6, 210)
        Me.btn_req_mkt_prc.Margin = New System.Windows.Forms.Padding(2)
        Me.btn_req_mkt_prc.Name = "btn_req_mkt_prc"
        Me.btn_req_mkt_prc.Size = New System.Drawing.Size(87, 21)
        Me.btn_req_mkt_prc.TabIndex = 2
        Me.btn_req_mkt_prc.Text = "Req Mkt Price"
        Me.btn_req_mkt_prc.UseVisualStyleBackColor = True
        '
        'lbl_connection
        '
        Me.lbl_connection.AutoSize = True
        Me.lbl_connection.Location = New System.Drawing.Point(3, 16)
        Me.lbl_connection.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lbl_connection.Name = "lbl_connection"
        Me.lbl_connection.Size = New System.Drawing.Size(61, 13)
        Me.lbl_connection.TabIndex = 3
        Me.lbl_connection.Text = "Connection"
        '
        'btn_req_TWS_opn_ords
        '
        Me.btn_req_TWS_opn_ords.Location = New System.Drawing.Point(6, 114)
        Me.btn_req_TWS_opn_ords.Name = "btn_req_TWS_opn_ords"
        Me.btn_req_TWS_opn_ords.Size = New System.Drawing.Size(117, 23)
        Me.btn_req_TWS_opn_ords.TabIndex = 4
        Me.btn_req_TWS_opn_ords.Text = "Req TWS Open Ords"
        Me.btn_req_TWS_opn_ords.UseVisualStyleBackColor = True
        '
        'btn_req_opn_ords
        '
        Me.btn_req_opn_ords.Location = New System.Drawing.Point(6, 143)
        Me.btn_req_opn_ords.Name = "btn_req_opn_ords"
        Me.btn_req_opn_ords.Size = New System.Drawing.Size(117, 23)
        Me.btn_req_opn_ords.TabIndex = 5
        Me.btn_req_opn_ords.Text = "Req Open Ords"
        Me.btn_req_opn_ords.UseVisualStyleBackColor = True
        '
        'tbx_quantity
        '
        Me.tbx_quantity.Location = New System.Drawing.Point(294, 32)
        Me.tbx_quantity.Name = "tbx_quantity"
        Me.tbx_quantity.Size = New System.Drawing.Size(35, 20)
        Me.tbx_quantity.TabIndex = 6
        '
        'txb_price
        '
        Me.txb_price.Location = New System.Drawing.Point(337, 32)
        Me.txb_price.Name = "txb_price"
        Me.txb_price.Size = New System.Drawing.Size(35, 20)
        Me.txb_price.TabIndex = 7
        '
        'cbx_action
        '
        Me.cbx_action.FormattingEnabled = True
        Me.cbx_action.Items.AddRange(New Object() {"BUY", "SELL"})
        Me.cbx_action.Location = New System.Drawing.Point(234, 31)
        Me.cbx_action.Name = "cbx_action"
        Me.cbx_action.Size = New System.Drawing.Size(50, 21)
        Me.cbx_action.TabIndex = 8
        '
        'btn_order
        '
        Me.btn_order.Location = New System.Drawing.Point(265, 59)
        Me.btn_order.Name = "btn_order"
        Me.btn_order.Size = New System.Drawing.Size(75, 23)
        Me.btn_order.TabIndex = 9
        Me.btn_order.Text = "Execute"
        Me.btn_order.UseVisualStyleBackColor = True
        '
        'lbl_action
        '
        Me.lbl_action.AutoSize = True
        Me.lbl_action.Location = New System.Drawing.Point(245, 16)
        Me.lbl_action.Name = "lbl_action"
        Me.lbl_action.Size = New System.Drawing.Size(37, 13)
        Me.lbl_action.TabIndex = 10
        Me.lbl_action.Text = "Action"
        '
        'lbl_quantity
        '
        Me.lbl_quantity.AutoSize = True
        Me.lbl_quantity.Location = New System.Drawing.Point(290, 16)
        Me.lbl_quantity.Name = "lbl_quantity"
        Me.lbl_quantity.Size = New System.Drawing.Size(46, 13)
        Me.lbl_quantity.TabIndex = 11
        Me.lbl_quantity.Text = "Quantity"
        '
        'lbl_price
        '
        Me.lbl_price.AutoSize = True
        Me.lbl_price.Location = New System.Drawing.Point(342, 16)
        Me.lbl_price.Name = "lbl_price"
        Me.lbl_price.Size = New System.Drawing.Size(31, 13)
        Me.lbl_price.TabIndex = 12
        Me.lbl_price.Text = "Price"
        '
        'btn_cancel
        '
        Me.btn_cancel.Location = New System.Drawing.Point(317, 88)
        Me.btn_cancel.Name = "btn_cancel"
        Me.btn_cancel.Size = New System.Drawing.Size(55, 23)
        Me.btn_cancel.TabIndex = 13
        Me.btn_cancel.Text = "Cancel"
        Me.btn_cancel.UseVisualStyleBackColor = True
        '
        'tbx_Order_ID
        '
        Me.tbx_Order_ID.Location = New System.Drawing.Point(276, 90)
        Me.tbx_Order_ID.Name = "tbx_Order_ID"
        Me.tbx_Order_ID.Size = New System.Drawing.Size(34, 20)
        Me.tbx_Order_ID.TabIndex = 14
        '
        'lbl_Order_ID
        '
        Me.lbl_Order_ID.AutoSize = True
        Me.lbl_Order_ID.Location = New System.Drawing.Point(220, 93)
        Me.lbl_Order_ID.Name = "lbl_Order_ID"
        Me.lbl_Order_ID.Size = New System.Drawing.Size(50, 13)
        Me.lbl_Order_ID.TabIndex = 15
        Me.lbl_Order_ID.Text = "Order ID:"
        '
        'btn_req_ord_ID
        '
        Me.btn_req_ord_ID.Location = New System.Drawing.Point(223, 143)
        Me.btn_req_ord_ID.Name = "btn_req_ord_ID"
        Me.btn_req_ord_ID.Size = New System.Drawing.Size(80, 23)
        Me.btn_req_ord_ID.TabIndex = 16
        Me.btn_req_ord_ID.Text = "Req Order ID"
        Me.btn_req_ord_ID.UseVisualStyleBackColor = True
        '
        'btn_trade_last_tick
        '
        Me.btn_trade_last_tick.Location = New System.Drawing.Point(6, 280)
        Me.btn_trade_last_tick.Name = "btn_trade_last_tick"
        Me.btn_trade_last_tick.Size = New System.Drawing.Size(75, 23)
        Me.btn_trade_last_tick.TabIndex = 17
        Me.btn_trade_last_tick.Text = "Trade - Last Tick"
        Me.btn_trade_last_tick.UseVisualStyleBackColor = True
        '
        'lbl_requesting_open_orders
        '
        Me.lbl_requesting_open_orders.AutoSize = True
        Me.lbl_requesting_open_orders.Location = New System.Drawing.Point(3, 88)
        Me.lbl_requesting_open_orders.Name = "lbl_requesting_open_orders"
        Me.lbl_requesting_open_orders.Size = New System.Drawing.Size(124, 13)
        Me.lbl_requesting_open_orders.TabIndex = 18
        Me.lbl_requesting_open_orders.Text = "Requesting Open Orders"
        '
        'lbl_downloading_tick_data
        '
        Me.lbl_downloading_tick_data.AutoSize = True
        Me.lbl_downloading_tick_data.Location = New System.Drawing.Point(6, 185)
        Me.lbl_downloading_tick_data.Name = "lbl_downloading_tick_data"
        Me.lbl_downloading_tick_data.Size = New System.Drawing.Size(119, 13)
        Me.lbl_downloading_tick_data.TabIndex = 19
        Me.lbl_downloading_tick_data.Text = "Downlaoding Tick Data"
        '
        'lbl_auto_trading
        '
        Me.lbl_auto_trading.AutoSize = True
        Me.lbl_auto_trading.Location = New System.Drawing.Point(6, 254)
        Me.lbl_auto_trading.Name = "lbl_auto_trading"
        Me.lbl_auto_trading.Size = New System.Drawing.Size(68, 13)
        Me.lbl_auto_trading.TabIndex = 20
        Me.lbl_auto_trading.Text = "Auto Trading"
        '
        'lbl_candle_resolution
        '
        Me.lbl_candle_resolution.AutoSize = True
        Me.lbl_candle_resolution.Location = New System.Drawing.Point(99, 254)
        Me.lbl_candle_resolution.Name = "lbl_candle_resolution"
        Me.lbl_candle_resolution.Size = New System.Drawing.Size(93, 13)
        Me.lbl_candle_resolution.TabIndex = 21
        Me.lbl_candle_resolution.Text = "Candle Resolution"
        '
        'tbx_candle_resolution
        '
        Me.tbx_candle_resolution.Location = New System.Drawing.Point(123, 270)
        Me.tbx_candle_resolution.Name = "tbx_candle_resolution"
        Me.tbx_candle_resolution.Size = New System.Drawing.Size(31, 20)
        Me.tbx_candle_resolution.TabIndex = 22
        '
        'lbl_for_bollinger_bands
        '
        Me.lbl_for_bollinger_bands.AutoSize = True
        Me.lbl_for_bollinger_bands.Location = New System.Drawing.Point(245, 218)
        Me.lbl_for_bollinger_bands.Name = "lbl_for_bollinger_bands"
        Me.lbl_for_bollinger_bands.Size = New System.Drawing.Size(98, 13)
        Me.lbl_for_bollinger_bands.TabIndex = 23
        Me.lbl_for_bollinger_bands.Text = "For Bollinger Bands"
        '
        'tbx_std_dev
        '
        Me.tbx_std_dev.Location = New System.Drawing.Point(248, 270)
        Me.tbx_std_dev.Name = "tbx_std_dev"
        Me.tbx_std_dev.Size = New System.Drawing.Size(22, 20)
        Me.tbx_std_dev.TabIndex = 24
        '
        'lbl_std_dev
        '
        Me.lbl_std_dev.AutoSize = True
        Me.lbl_std_dev.Location = New System.Drawing.Point(234, 253)
        Me.lbl_std_dev.Name = "lbl_std_dev"
        Me.lbl_std_dev.Size = New System.Drawing.Size(46, 13)
        Me.lbl_std_dev.TabIndex = 25
        Me.lbl_std_dev.Text = "Std Dev"
        '
        'lbl_min_num_of_candles
        '
        Me.lbl_min_num_of_candles.AutoSize = True
        Me.lbl_min_num_of_candles.Location = New System.Drawing.Point(99, 293)
        Me.lbl_min_num_of_candles.Name = "lbl_min_num_of_candles"
        Me.lbl_min_num_of_candles.Size = New System.Drawing.Size(86, 13)
        Me.lbl_min_num_of_candles.TabIndex = 26
        Me.lbl_min_num_of_candles.Text = "Min # of candles"
        '
        'tbx_min_num_of_candles
        '
        Me.tbx_min_num_of_candles.Location = New System.Drawing.Point(123, 309)
        Me.tbx_min_num_of_candles.Name = "tbx_min_num_of_candles"
        Me.tbx_min_num_of_candles.Size = New System.Drawing.Size(16, 20)
        Me.tbx_min_num_of_candles.TabIndex = 27
        '
        'lbl_lookbck_for_calcs
        '
        Me.lbl_lookbck_for_calcs.AutoSize = True
        Me.lbl_lookbck_for_calcs.Location = New System.Drawing.Point(281, 254)
        Me.lbl_lookbck_for_calcs.Name = "lbl_lookbck_for_calcs"
        Me.lbl_lookbck_for_calcs.Size = New System.Drawing.Size(92, 13)
        Me.lbl_lookbck_for_calcs.TabIndex = 28
        Me.lbl_lookbck_for_calcs.Text = "Lookbck for calcs"
        '
        'tbx_lookback_for_calcs
        '
        Me.tbx_lookback_for_calcs.Location = New System.Drawing.Point(294, 270)
        Me.tbx_lookback_for_calcs.Name = "tbx_lookback_for_calcs"
        Me.tbx_lookback_for_calcs.Size = New System.Drawing.Size(62, 20)
        Me.tbx_lookback_for_calcs.TabIndex = 29
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 345)
        Me.Controls.Add(Me.tbx_lookback_for_calcs)
        Me.Controls.Add(Me.lbl_lookbck_for_calcs)
        Me.Controls.Add(Me.tbx_min_num_of_candles)
        Me.Controls.Add(Me.lbl_min_num_of_candles)
        Me.Controls.Add(Me.lbl_std_dev)
        Me.Controls.Add(Me.tbx_std_dev)
        Me.Controls.Add(Me.lbl_for_bollinger_bands)
        Me.Controls.Add(Me.tbx_candle_resolution)
        Me.Controls.Add(Me.lbl_candle_resolution)
        Me.Controls.Add(Me.lbl_auto_trading)
        Me.Controls.Add(Me.lbl_downloading_tick_data)
        Me.Controls.Add(Me.lbl_requesting_open_orders)
        Me.Controls.Add(Me.btn_trade_last_tick)
        Me.Controls.Add(Me.btn_req_ord_ID)
        Me.Controls.Add(Me.lbl_Order_ID)
        Me.Controls.Add(Me.tbx_Order_ID)
        Me.Controls.Add(Me.btn_cancel)
        Me.Controls.Add(Me.lbl_price)
        Me.Controls.Add(Me.lbl_quantity)
        Me.Controls.Add(Me.lbl_action)
        Me.Controls.Add(Me.btn_order)
        Me.Controls.Add(Me.cbx_action)
        Me.Controls.Add(Me.txb_price)
        Me.Controls.Add(Me.tbx_quantity)
        Me.Controls.Add(Me.btn_req_opn_ords)
        Me.Controls.Add(Me.btn_req_TWS_opn_ords)
        Me.Controls.Add(Me.lbl_connection)
        Me.Controls.Add(Me.btn_req_mkt_prc)
        Me.Controls.Add(Me.btn_req_time)
        Me.Controls.Add(Me.btn_cnct)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btn_cnct As Button
    Friend WithEvents btn_req_time As Button
    Friend WithEvents btn_req_mkt_prc As Button
    Friend WithEvents lbl_connection As Label
    Friend WithEvents btn_req_TWS_opn_ords As Button
    Friend WithEvents btn_req_opn_ords As Button
    Friend WithEvents tbx_quantity As TextBox
    Friend WithEvents txb_price As TextBox
    Friend WithEvents cbx_action As ComboBox
    Friend WithEvents btn_order As Button
    Friend WithEvents lbl_action As Label
    Friend WithEvents lbl_quantity As Label
    Friend WithEvents lbl_price As Label
    Friend WithEvents btn_cancel As Button
    Friend WithEvents tbx_Order_ID As TextBox
    Friend WithEvents lbl_Order_ID As Label
    Friend WithEvents btn_req_ord_ID As Button
    Friend WithEvents btn_trade_last_tick As Button
    Friend WithEvents lbl_requesting_open_orders As Label
    Friend WithEvents lbl_downloading_tick_data As Label
    Friend WithEvents lbl_auto_trading As Label
    Friend WithEvents lbl_candle_resolution As Label
    Friend WithEvents tbx_candle_resolution As TextBox
    Friend WithEvents lbl_for_bollinger_bands As Label
    Friend WithEvents tbx_std_dev As TextBox
    Friend WithEvents lbl_std_dev As Label
    Friend WithEvents lbl_min_num_of_candles As Label
    Friend WithEvents tbx_min_num_of_candles As TextBox
    Friend WithEvents lbl_lookbck_for_calcs As Label
    Friend WithEvents tbx_lookback_for_calcs As TextBox
End Class
