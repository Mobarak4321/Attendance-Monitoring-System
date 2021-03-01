Imports System.Data.OleDb

Public Class Login

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Connection()

        cnn.Open()

        cmd = New OleDbCommand("SELECT * FROM userTbl WHERE [Username] ='" & TextBox1.Text.Trim & "' AND [Userpassword]='" & TextBox2.Text.Trim & "' ", cnn)

        reader = cmd.ExecuteReader

        Dim userFound As Boolean = False

        Dim Username As String = ""
        Dim UserPassword As String = ""

        While reader.Read

            userFound = True

            Username = reader("Username").ToString
            UserPassword = reader("Userpassword").ToString

        End While

        If userFound = True Then

            MsgBox("Login complete")

            Home.Show()
            Me.Hide()

            TextBox1.Clear()
            TextBox2.Clear()

            cnn.Close()

        Else
            cnn.Close()
            MsgBox("Incorrect Credentials")
            TextBox1.Text = "Username"
            TextBox2.Text = "Password"

        End If
        cnn.Close()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Application.Exit()
    End Sub

    Private Sub TextBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox2.MouseClick
        TextBox2.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        RePassword.ShowDialog()

    End Sub
End Class