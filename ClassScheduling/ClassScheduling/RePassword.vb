Imports System.Data.OleDb

Public Class RePassword

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim existing As Boolean
        Try
            Call Connection()


            sql = "SELECT * FROM Code WHERE Code = '" & TextBox1.Text & "'"
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            If reader.Read Then
                If reader.HasRows Then


                    existing = True
                Else
                    existing = False
                End If
            End If
        Catch ex As Exception

        Finally
            cnn.Close()

            If existing = True Then

                TextBox2.Enabled = True
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                Button2.Enabled = True

            Else
                MsgBox("The Code:" & TextBox1.Text & "is incorrect")

            End If

        End Try


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Dispose()
        Login.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Connection()
        Dim cmd2 As New OleDbCommand

        If TextBox3.Text = TextBox4.Text then

        Try
            sql = "DELETE * FROM UserTbl"
            sql = "INSERT INTO UserTbl (Username, Password) VALUES ('" & TextBox2.Text & "','" & TextBox3.Text & "')"
            cmd = New OleDbCommand(sql, cnn)
            cmd2 = New OleDbCommand(sql, cnn)
            cnn.Open()

            cmd2.ExecuteNonQuery()
            cmd.ExecuteNonQuery()

            MsgBox("Successfully Created")

                Me.Dispose()
                Login.ShowDialog()

                cnn.Close()
        Catch ex As Exception

            End Try

        Else
            MsgBox("Password mismatch")
        End If
    End Sub
End Class