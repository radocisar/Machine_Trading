Imports System.IO

Public Class Create_new_folder_for_cntrt

    Public Shared path As String = "E:\Test_FX_Ticks_Download\" & Functions.cntrt.Symbol

    Public Shared Function file_path() As String

        Directory.CreateDirectory(path)

        file_path = path & "\"

    End Function

End Class
