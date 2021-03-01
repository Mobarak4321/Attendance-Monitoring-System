Imports System.Data.OleDb

Module SystemCore
    Public cnn As New OleDbConnection
    Public cmd As New OleDbCommand
    Public da As OleDbDataAdapter
    Public reader As OleDbDataReader
    Public conn_status As String
    Public sql As String
    Public Schedule_Subject, Schedule_Faculty, Schedule_Day, Schedule_TimeIn, Schedule_TimeOut, Schedule_Room, Schedule_Section As String
    Public IDRoom, RoomInfo As String


    Public Sub Connection()

        Try

            cnn.ConnectionString = "provider=Microsoft.JET.OLEDB.4.0; Data Source=" & Application.StartupPath & "\ScheduleDb.mdb"
            cmd.Connection = cnn
            cnn.Open()

        Catch ex As Exception

            MsgBox("Critical Error the System will shutdown error found:" & ex.Message & " Please Contact System Administrator", MsgBoxStyle.Critical)
            conn_status = "Offline"
            cnn.Close()
            End

        Finally

            cnn.Close()
            ' conn_status = "Online"
            ' MsgBox(conn_status)

        End Try

    End Sub

End Module
