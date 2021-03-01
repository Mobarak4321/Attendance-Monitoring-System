Imports System.Data.OleDb

Public Class StrandManagement

    Sub Load_Section()

        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM StrandTbl "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("StrandCode").ToString)
                x.SubItems.Add(reader("StrandDescription").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub SectionManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Load_Section()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM StrandTbl WHERE StrandCode like '%" & TextBox1.Text & "%' OR StrandDescrition like '%" & TextBox1.Text & "%'"
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("StrandCode").ToString)
                x.SubItems.Add(reader("StrandDescription").ToString)
                ListView1.Items.Add(x)

            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Home.switchPanel(StrandAdd)
        Me.Dispose()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class