Imports System.Data.OleDb

Public Class RoomAdd

    Sub RoomAdd()
        If TextBox1.Text.Trim <> "" Then


            Dim existing As Boolean

            Call Connection()

            Try
                sql = ("SELECT * FROM RoomTbl WHERE Room = '" & TextBox1.Text.Trim + "-" + TextBox2.Text & "' ")
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
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally

                cnn.Close()

                If existing = False Then

                    Try

                        sql = ("INSERT INTO RoomTbl (Room) VALUES ('" & TextBox1.Text.Trim + "-" + TextBox2.Text & "')")
                        cmd = New OleDbCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()
                        cnn.Close()

                        MsgBox("Successfully Added")
                        Home.switchPanel(FacultyManagement)

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        cnn.Close()
                        cmd.Dispose()

                    End Try

                Else

                    MsgBox("already registered")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RoomAdd()
    End Sub
End Class