Imports System.Data.OleDb
Imports System.Drawing.Printing
Public Class ScheduleManagement

    Dim H As Integer
    Dim LineNumber As Integer
    Dim Lineperpage As Integer
    Dim I_Start As Integer
    Dim I_Counter As Integer

    Sub Load_Schedule()
        Dim time As DateTime

        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM ScheduleTbl ORDER BY TimeIn "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem


            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("Subject").ToString)
                x.SubItems.Add(reader("Faculty").ToString)
                x.SubItems.Add(reader("Day").ToString)
                x.SubItems.Add(DateTime.Parse(reader("TimeIn").ToString()).ToString("hh:mm tt"))
                x.SubItems.Add(DateTime.Parse(reader("TimeOut").ToString()).ToString("hh:mm tt"))
                x.SubItems.Add(reader("Room").ToString)
                x.SubItems.Add(reader("Strand").ToString)
                x.SubItems.Add(reader("Grade").ToString)
                x.SubItems.Add(reader("Section").ToString)

                ListView1.Items.Add(x)
            Loop
            time = Format(x.SubItems(4).Text, "hh:mm: tt")
        Catch ex As Exception

            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub ScheduleManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            Lineperpage = 30
            Load_Schedule()
            Call Connection()


            Dim command As New OleDbCommand("SELECT * FROM RoomTbl ", cnn)
            Dim adapter As New OleDbDataAdapter(command)

            Dim Table As New DataTable

            adapter.Fill(Table)

            ComboBox2.DataSource = Table

            ComboBox2.DisplayMember = "Room"
            ComboBox2.ValueMember = "ID"

            cnn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)


        End Try

    End Sub

    Private Sub TextBox1_Enter(sender As Object, e As EventArgs) Handles TextBox1.Enter
        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM ScheduleTbl WHERE Subject = '" & TextBox1.Text & "'"
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("Subject").ToString)
                x.SubItems.Add(reader("Faculty").ToString)
                x.SubItems.Add(reader("Day").ToString)
                x.SubItems.Add(reader("TimeIn").ToString)
                x.SubItems.Add(reader("TimeOut").ToString)
                x.SubItems.Add(reader("Room").ToString)
                x.SubItems.Add(reader("Strand").ToString)
                x.SubItems.Add(reader("Grade").ToString)
                x.SubItems.Add(reader("Section").ToString)
                ListView1.Items.Add(x)
            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM ScheduleTbl WHERE Day= '" & ComboBox1.Text & "' "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("Subject").ToString)
                x.SubItems.Add(reader("Faculty").ToString)
                x.SubItems.Add(reader("Day").ToString)
                x.SubItems.Add(reader("TimeIn").ToString)
                x.SubItems.Add(reader("TimeOut").ToString)
                x.SubItems.Add(reader("Room").ToString)
                x.SubItems.Add(reader("Strand").ToString)
                x.SubItems.Add(reader("Grade").ToString)
                x.SubItems.Add(reader("Section").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM ScheduleTbl WHERE Room = '" & ComboBox2.Text & "'"
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("Subject").ToString)
                x.SubItems.Add(reader("Faculty").ToString)
                x.SubItems.Add(reader("Day").ToString)
                x.SubItems.Add(reader("TimeIn").ToString)
                x.SubItems.Add(reader("TimeOut").ToString)
                x.SubItems.Add(reader("Room").ToString)
                x.SubItems.Add(reader("Strand").ToString)
                x.SubItems.Add(reader("Grade").ToString)
                x.SubItems.Add(reader("Section").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        Dim ID As String = ListView1.SelectedItems(0).SubItems(0).Text
        Dim Subject As String = ListView1.SelectedItems(0).SubItems(1).Text
        Dim Faculty As String = ListView1.SelectedItems(0).SubItems(2).Text
        Dim Day As String = ListView1.SelectedItems(0).SubItems(3).Text
        Dim TimeIn As String = ListView1.SelectedItems(0).SubItems(4).Text
        Dim TimeOut As String = ListView1.SelectedItems(0).SubItems(5).Text
        Dim Room As String = ListView1.SelectedItems(0).SubItems(6).Text
        Dim Grade As String = ListView1.SelectedItems(0).SubItems(7).Text
        Dim Section As String = ListView1.SelectedItems(0).SubItems(8).Text


        ScheduleUpdate.Label9.Text = ID
        ScheduleUpdate.ComboBox1.Text = Subject
        ScheduleUpdate.ComboBox2.Text = Faculty
        ScheduleUpdate.ComboBox3.Text = Day
        ScheduleUpdate.DateTimePicker1.Text = TimeIn
        ScheduleUpdate.DateTimePicker2.Text = TimeOut
        ScheduleUpdate.ComboBox4.Text = Room
        ScheduleUpdate.ComboBox5.Text = Grade
        ScheduleUpdate.ComboBox6.Text = Section

    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Home.switchPanel(ScheduleUpdate)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM ScheduleTbl WHERE Day= '" & ComboBox1.Text & "' AND Room= '" & ComboBox2.Text & "' "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("Subject").ToString)
                x.SubItems.Add(reader("Faculty").ToString)
                x.SubItems.Add(reader("Day").ToString)
                x.SubItems.Add(reader("TimeIn").ToString)
                x.SubItems.Add(reader("TimeOut").ToString)
                x.SubItems.Add(reader("Room").ToString)
                x.SubItems.Add(reader("Strand").ToString)
                x.SubItems.Add(reader("Grade").ToString)
                x.SubItems.Add(reader("Section").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings

            Dim New_Paper As New PaperSize("", 800, 800)

            New_Paper.PaperName = PaperKind.Custom
            Dim New_Margin As New Margins
            New_Margin.Left = 0
            New_Margin.Top = 50

            With PrintDocument1
                .DefaultPageSettings.PaperSize = New_Paper
                .DefaultPageSettings.Margins = New_Margin

                .PrinterSettings.DefaultPageSettings.Landscape = False
                .Print()
            End With
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage

        H = 50
        e.Graphics.DrawString("Class Schedule", New Drawing.Font("Times New Roman", 12), Brushes.Blue, 320, H)

        H += 30
        e.Graphics.DrawString(Label1.Text, New Drawing.Font("Times New Roman", 12), Brushes.Black, 20, H)

        H += 20
        Dim NN As Integer = H
        e.Graphics.DrawLine(Pens.Red, 7, NN, 800, NN)

        H += 30

        e.Graphics.DrawString("Subject", New Drawing.Font("Times New Roman", 12), Brushes.Black, 10, H)
        e.Graphics.DrawString("Room", New Drawing.Font("Times New Roman", 12), Brushes.Black, 120, H)
        e.Graphics.DrawString("Day", New Drawing.Font("Times New Roman", 12), Brushes.Black, 420, H)


        H += 50

        'For Each Itm As ListViewItem In ListView1.Items
        For Me.I_Counter = I_Start To ListView1.Items.Count - 1

            'e.Graphics.DrawString(Itm.Text, New Drawing.Font("Times New Roman", 12), Brushes.Black, 50, H)
            e.Graphics.DrawString(ListView1.Items(I_Counter).SubItems(1).Text, New Drawing.Font("Times New Roman", 12), Brushes.Black, 10, H)
            e.Graphics.DrawString(ListView1.Items(I_Counter).SubItems(6).Text, New Drawing.Font("Times New Roman", 12), Brushes.Black, 120, H)
            e.Graphics.DrawString(ListView1.Items(I_Counter).SubItems(3).Text, New Drawing.Font("Times New Roman", 12), Brushes.Black, 420, H)

            H += 20

            LineNumber += 1
            If LineNumber = Lineperpage Then
                LineNumber = 0
                I_Start = I_Counter + 1
                H = 50

                e.HasMorePages = True
                Exit For
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Home.switchPanel(ScheduleAdd)
        ScheduleAdd.Timer1.Start()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call Connection()

        Try
            sql = ("DELETE FROM ScheduleTbl")
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            MsgBox("All the Schedule that already register is will be deleted")
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("error " & ex.Message & " " & ex.StackTrace)

        Finally
            cnn.Close()
        End Try
    End Sub
End Class