Public Class Bollinger_Band_Strategy
    Private WithEvents Candles_creation_data_processing_event As New Candles_Creation_Data_Processing

    Public upper_band As Double
    Public lower_band As Double
    Public middle_band As Double
    Public Upper_Lower_Band_Span As Double

    Sub Bollinger_Band_Trading_Strategy(candle_arr() As Candle, open_price As Double, high_price As Double, low_price As Double, last_price As Double) Handles Candles_creation_data_processing_event.start_trading_strategy

        'upper_band =
        'lower_band =
        'middle_band =
        'Upper_Lower_Band_Span = 



    End Sub

End Class
