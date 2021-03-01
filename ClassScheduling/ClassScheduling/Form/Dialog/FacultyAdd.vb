Imports System.Data.OleDb

Public Class FacultyAdd

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FacultyAdd()
    End Sub

    Sub FacultyAdd()
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "" Or TextBox3.Text.Trim <> "" Then

            Dim d As DateTime = Now
            Debug.WriteLine(d.ToLongDateString)

            Dim existing As Boolean

            Call Connection()

            Try
                sql = "SELECT * FROM FacultyTbl WHERE Lastname= '" & TextBox1.Text.Trim & "' AND Firstname= '" & TextBox2.Text & "' AND Middlename= '" & TextBox3.Text & "' "
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

                        sql = ("INSERT INTO FacultyTbl (Lastname, Firstname, Middlename) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "','" & TextBox3.Text.Trim & "')")
                        cmd = New OleDbCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Added")
                        Home.switchPanel(FacultyManagement)

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
End Class