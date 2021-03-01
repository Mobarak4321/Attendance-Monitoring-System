Imports System.Data.OleDb

Public Class SubjectManagement

    Sub Load_Subject()

        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM SubjectTbl "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("SubjectID").ToString)
                x.SubItems.Add(reader("SubjectCode").ToString)
                x.SubItems.Add(reader("SubjectDescription").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub SubjectManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        load_subject()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Home.switchPanel(SubjectAdd)
    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick

    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Home.switchPanel(SubjectUpdate)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM SubjectTbl WHERE SubjectCode = '" & TextBox1.Text & "' OR SubjectDescription = '" & TextBox1.Text & "'"
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("SubjectID").ToString)
                x.SubItems.Add(reader("SubjectCode").ToString)
                x.SubItems.Add(reader("SubjectDescription").ToString)
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