Imports System.Data.OleDb

Public Class ScheduleAdd

    Private Sub ScheduleAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            ScheduleManagement.Load_Schedule()
            Call Connection()


            Dim command As New OleDbCommand("SELECT * FROM RoomTbl ", cnn)
            Dim command2 As New OleDbCommand("SELECT * FROM SubjectTbl ", cnn)
            Dim command3 As New OleDbCommand("SELECT * FROM SectionTbl ", cnn)
            Dim command4 As New OleDbCommand("SELECT * FROM FacultyTbl ", cnn)
            Dim command5 As New OleDbCommand("SELECT * FROM StrandTbl ", cnn)

            Dim adapter As New OleDbDataAdapter(command)
            Dim adapter2 As New OleDbDataAdapter(command2)
            Dim adapter3 As New OleDbDataAdapter(command3)
            Dim adapter4 As New OleDbDataAdapter(command4)
            Dim adapter5 As New OleDbDataAdapter(command5)

            Dim Table As New DataTable
            Dim Table2 As New DataTable
            Dim Table3 As New DataTable
            Dim Table4 As New DataTable
            Dim Table5 As New DataTable

            adapter.Fill(Table)
            adapter2.Fill(Table2)
            adapter3.Fill(Table3)
            adapter4.Fill(Table4)
            adapter5.Fill(Table5)

            ComboBox5.DataSource = Table
            ComboBox1.DataSource = Table2
            ComboBox6.DataSource = Table3
            ComboBox2.DataSource = Table4
            ComboBox7.DataSource = Table5

            ComboBox5.DisplayMember = "Room"
            ComboBox5.ValueMember = "ID"

            ComboBox2.DisplayMember = "Lastname"
            ComboBox2.ValueMember = "FacultyID"

            ComboBox1.DisplayMember = "SubjectCode"
            ComboBox1.ValueMember = "SubjectID"

            ComboBox6.DisplayMember = "Section"
            ComboBox6.ValueMember = "SectionID"

            ComboBox7.DisplayMember = "Section"
            ComboBox7.ValueMember = "StrandCode"
            cnn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)


        End Try
    End Sub

    Sub ScheduleAdd()



        If ComboBox1.Text.Trim <> "" Or ComboBox2.Text.Trim <> "" Or ComboBox3.Text.Trim <> "" Then

            Dim existing As Boolean

            Dim Day As String = ComboBox3.Text
            Dim Room As String = ComboBox5.Text
            Dim TimeIn As String = DateTimePicker1.Text
            Dim TimeOut As String = DateTimePicker2.Text

            Dim FirstTimeIn As String = DateTimePicker1.Text
            Dim FirstTimeOut As String = DateTimePicker2.Text
            Dim SecondTimeIn As String = DateTimePicker1.Text
            Dim SecondTimeOut As String = DateTimePicker2.Text
            Dim Subject As String

            Call Connection()

            Try
                sql = "SELECT * FROM ScheduleTbl WHERE  (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeIn BETWEEN ('" & DateTimePicker1.Text & "') AND ('" & DateTimePicker2.Text & "')) OR (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeOut BETWEEN ('" & DateTimePicker2.Text & "') AND ('" & DateTimePicker1.Text & "')) OR (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeIn < '" & DateTimePicker1.Text & "' AND TimeOut > '" & DateTimePicker2.Text & "')"

                'sql = "SELECT * FROM ScheduleTbl WHERE  (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeIn BETWEEN ('" & DateTimePicker1.Text & "') AND ('" & DateTimePicker2.Text & "')) OR (Day = '" & Day & "' AND Room = '" & Room & "' AND TimeOut BETWEEN ('" & DateTimePicker2.Text & "') AND ('" & DateTimePicker1.Text & "'))"

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

                        Dim str As String
                        str = "INSERT INTO ScheduleTbl ([Subject], [Faculty], [Day], [TimeIn], [TimeOut],[Room],[Strand],[Grade],[Section]) Values (?,?,?,?,?,?,?,?,?)"
                        Dim cmd As OleDbCommand = New OleDbCommand(str, cnn)
                        cnn.Open()

                        cmd.Parameters.Add(New OleDbParameter("Subject", CType(ComboBox1.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("Faculty", CType(ComboBox2.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("Day", CType(ComboBox3.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("TimeIn", CType(DateTimePicker1.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("TimeOut", CType(DateTimePicker2.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("Room", CType(ComboBox5.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("Strand", CType(ComboBox7.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("Grade", CType(ComboBox4.Text, String)))
                        cmd.Parameters.Add(New OleDbParameter("Section", CType(ComboBox6.Text, String)))

                        cmd.ExecuteNonQuery()

                        MsgBox("Successfully Added")
                        cnn.Close()
                        Me.Close()

                    Catch ex As SystemException

                        MsgBox("error " & ex.Message & " " & ex.StackTrace)

                        cnn.Close()

                    Finally

                        'Call Clear()

                    End Try

                Else
                    MsgBox("Schedule is already registered or conflict")
                    cnn.Close()
                End If

            End Try

        Else
            MsgBox("fill all the blanks")

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ScheduleAdd()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
        Home.switchPanel(ScheduleManagement)
    End Sub
End Class