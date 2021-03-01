Imports System.Data.OleDb

Public Class ScheduleUpdate

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Update22()
    End Sub

    Sub Update22()

        If ComboBox1.Text.Trim <> "" Or ComboBox2.Text.Trim <> "" Or ComboBox3.Text.Trim <> "" Or ComboBox4.Text.Trim <> "" Or ComboBox5.Text.Trim <> "" Or ComboBox6.Text.Trim <> "" Then

            Dim existing As Boolean

            Dim Day As String = ComboBox3.Text
            Dim Room As String = ComboBox5.Text
            Dim TimeIn As String = DateTimePicker1.Text
            Dim TimeOut As String = DateTimePicker2.Text

            Dim FirstTimeIn As String = DateTimePicker1.Text
            Dim FirstTimeOut As String = DateTimePicker2.Text
            Dim SecondTimeIn As String = DateTimePicker1.Text
            Dim SecondTimeOut As String = DateTimePicker2.Text

            Call Connection()

            Try

                sql = "SELECT * FROM ScheduleTbl WHERE  (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeIn BETWEEN ('" & DateTimePicker1.Text & "') AND ('" & DateTimePicker2.Text & "')) OR (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeOut BETWEEN ('" & DateTimePicker2.Text & "') AND ('" & DateTimePicker1.Text & "')) OR (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeIn < '" & DateTimePicker1.Text & "' AND TimeOut > '" & DateTimePicker2.Text & "')"

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

                        sql = ("UPDATE ScheduleTbl SET [Subject] = '" & ComboBox1.Text.Trim & "', [Faculty] = '" & ComboBox2.Text.Trim & "', [Day] = '" & ComboBox3.Text.Trim & "', [TimeIn] = '" & DateTimePicker1.Text & "', [TimeOut] = '" & DateTimePicker2.Text & "', [Room] = '" & ComboBox4.Text.Trim & "', [Grade] = '" & ComboBox5.Text.Trim & "', [Section] = '" & ComboBox6.Text.Trim & "' WHERE ID= " & Label9.Text & "")
                        cmd = New OleDbCommand(sql, cnn)
                        cnn.Open()

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Updated")
                        cnn.Close()
                        cmd.Dispose()
                        Me.Dispose()
                        Home.switchPanel(ScheduleManagement)

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