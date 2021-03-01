Imports System.Data.OleDb

Public Class SubjectUpdate

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim <> "" Or TextBox2.Text.Trim <> "Then" Then

            Call Connection()

            Try
                sql = ("UPDATE SubjectTbl SET SubjectCode= '" & TextBox1.Text.Trim & "', SubjectDescription= '" & TextBox2.Text.Trim & "' WHERE ID= " & Label3.Text.Trim & "")
                cmd = New OleDbCommand(sql, cnn)
                cnn.Open()

                cmd.ExecuteNonQuery()

                MsgBox("Successfully Updated")
                cnn.Close()
                cmd.Dispose()
                Me.Dispose()
                Home.switchPanel(FacultyManagement)

            Catch ex As Exception
                MsgBox("error found ;" & ex.Message & ex.StackTrace)
            Finally
                cnn.Close()
            End Try

        End If
    End Sub
End Class