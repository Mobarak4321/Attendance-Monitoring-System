Imports System.Data.OleDb

Public Class RoomUpdate

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Then

            Call Connection()

            Try
                sql = ("UPDATE RoomTbl SET Room= '" & TextBox1.Text.Trim & "' WHERE ID= " & TextBox1.Text.Trim & "")
                cmd = New OleDbCommand(sql, cnn)
                cnn.Open()

                cmd.ExecuteNonQuery()

                MsgBox("Successfully Updated")
                cnn.Close()
                cmd.Dispose()
                Me.Dispose()
                Home.switchPanel(RoomManagement)

            Catch ex As Exception
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally
                cnn.Close()
            End Try

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
        Home.switchPanel(RoomManagement)
    End Sub
End Class