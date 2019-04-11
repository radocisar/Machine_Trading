Imports System.IO

Public Class Open_str_wrt

    Public Shared str_wrt As StreamWriter
    Public Shared str_wrt_1 As StreamWriter

    Public Sub Open_str_wrt_method()

        'Dim str_wrt As StreamWriter = New StreamWriter("C:\Raw_Data\Logging _Auto_Trading\log.txt", True)
        str_wrt = New StreamWriter("D:\Auto_Trading_Logs\log.txt", True)

    End Sub

    Public Sub Open_str_wrt_method_1()

        'Dim str_wrt As StreamWriter = New StreamWriter("C:\Raw_Data\Logging _Auto_Trading\log.txt", True)
        str_wrt_1 = New StreamWriter("E:\Test_FX_Ticks_Download\log.txt", True)

    End Sub

End Class