Imports System.Data.OleDb

Public Class StrandAdd

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        StarndAdd()
    End Sub

    Sub StarndAdd()
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Then

            Dim d As DateTime = Now
            Debug.WriteLine(d.ToLongDateString)

            Dim existing As Boolean

            Call Connection()

            Try
                sql = "SELECT * FROM StrandTbl WHERE StrandCode= '" & TextBox1.Text.Trim & "' AND StrandDescription= '" & TextBox2.Text & "' "
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

                        sql = ("INSERT INTO StrandTbl (StrandCode, StrandDescription) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "')")
                        cmd = New OleDbCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Added")
                        Home.switchPanel(StrandManagement)

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        cnn.Close()
                        cmd.Dispose()
                        'Call Clear()

                    End Try

                Else

                    MsgBox("already registered")

                End If

            End Try

        Else

            MsgBox("fill all the blanks")

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
        Home.switchPanel(StrandManagement)
    End Sub
End Class