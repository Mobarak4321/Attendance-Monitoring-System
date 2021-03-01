Imports System.Data.OleDb

Public Class SubjectAdd

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SubjectAdd()
    End Sub

    Sub SubjectAdd()
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> " Then" Then


            Dim existing As Boolean

            Call Connection()

            Try
                sql = "SELECT * FROM SubjectTbl WHERE SubjectCode= '" & TextBox1.Text.Trim & "' AND SubjectDescription= '" & TextBox2.Text & "' "
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

                        sql = ("INSERT INTO SubjectTbl (SubjectCode, SubjectDescription) VALUES ('" & TextBox1.Text.Trim & "', '" & TextBox2.Text.Trim & "')")
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
End Class