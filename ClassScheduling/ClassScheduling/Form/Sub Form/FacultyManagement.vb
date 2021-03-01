Imports System.Data.OleDb

Public Class FacultyManagement

    Sub Load_Faculty()

        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM FacultyTbl "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("FacultyID").ToString)
                x.SubItems.Add(reader("Lastname").ToString)
                x.SubItems.Add(reader("Firstname").ToString)
                x.SubItems.Add(reader("Middlename").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub
    Private Sub FacultyManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Load_Faculty()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Home.switchPanel(FacultyAdd)
    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        Dim ID As String = ListView1.SelectedItems(0).SubItems(0).Text
        Dim Lastname As String = ListView1.SelectedItems(0).SubItems(1).Text
        Dim Firstname As String = ListView1.SelectedItems(0).SubItems(2).Text
        Dim Middlename As String = ListView1.SelectedItems(0).SubItems(3).Text


        FacultyUpdate.Label4.Text = ID
        FacultyUpdate.TextBox1.Text = Lastname
        FacultyUpdate.TextBox2.Text = Firstname
        FacultyUpdate.TextBox3.Text = Middlename
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Home.switchPanel(FacultyUpdate)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM FacultyTbl WHERE Lastname = '" & TextBox1.Text & "' OR Firstname = '" & TextBox1.Text & "' OR Middlename = '" & TextBox1.Text & "' "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("FacultyID").ToString)
                x.SubItems.Add(reader("Lastname").ToString)
                x.SubItems.Add(reader("Firstname").ToString)
                x.SubItems.Add(reader("Middlename").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub
End Class