Public Class Bollinger_Band_Strategy
    Private WithEvents Candles_creation_data_processing_event As New Candles_Creation_Data_Processing

    Public upper_band As Double
    Public lower_band As Double
    Public middle_band As Double
    Public Upper_Lower_Band_Span As Double

    Sub Bollinger_Band_Trading_Strategy(BB_candle_arr(,) As Double, open_price As Double, high_price As Double, low_price As Double, last_price As Double) Handles Candles_creation_data_processing_event.start_trading_strategy

        ' Shift all values by one so that last candle can be updated with latest data (open, high, low, close)
        For n = 0 To Form1.candle_init_count - 2
            BB_candle_arr(n, 0) = BB_candle_arr(n + 1, 0)
            BB_candle_arr(n, 1) = BB_candle_arr(n + 1, 1)
            BB_candle_arr(n, 2) = BB_candle_arr(n + 1, 2)
            BB_candle_arr(n, 3) = BB_candle_arr(n + 1, 3)
        Next

        BB_candle_arr(Form1.candle_init_count - 1, 0) = open_price
        BB_candle_arr(Form1.candle_init_count - 1, 1) = high_price
        BB_candle_arr(Form1.candle_init_count - 1, 2) = low_price
        BB_candle_arr(Form1.candle_init_count - 1, 3) = last_price

#Region "Middle band"
        middle_band = calculate_mean(BB_candle_arr)
#End Region

#Region "Upper band"
        upper_band = middle_band + (calculate_std(middle_band, BB_candle_arr) * Form1.tbx_std_dev)
#End Region

#Region "Lower band"
        lower_band = middle_band - (calculate_std(middle_band, BB_candle_arr) * Form1.tbx_std_dev)
#End Region

#Region "Band span"
        Upper_Lower_Band_Span = upper_band - lower_band
#End Region

    End Sub

    Function calculate_mean(BB_candle_arr(,) As Double)
        Dim sum_of_close As Double

        For n = 0 To Form1.candle_init_count - 1
            sum_of_close = sum_of_close + BB_candle_arr(n, 3)
        Next

        middle_band = sum_of_close / BB_candle_arr.Length

        Return middle_band

    End Function

    Function calculate_std(middle_band As Double, BB_candle_arr(,) As Double)

        Dim value_and_mean_diff As Double
        Dim std As Double

        For n = 0 To Form1.candle_init_count - 1
            value_and_mean_diff = value_and_mean_diff + ((BB_candle_arr(n, 3) - middle_band) * (BB_candle_arr(n, 3) - middle_band))
        Next

        std = Math.Sqrt(value_and_mean_diff / BB_candle_arr.Length)

        Return std

    End Function
End Class
