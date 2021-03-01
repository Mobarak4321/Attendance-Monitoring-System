Imports System.Data.OleDb

Public Class RoomManagement

    Sub Load_Room()

        Call Connection()

        ListView1.Items.Clear()

        Try
            sql = "SELECT * FROM RoomTbl "
            cmd = New OleDbCommand(sql, cnn)
            cnn.Open()

            reader = cmd.ExecuteReader

            Dim x As ListViewItem

            Do While reader.Read = True

                x = New ListViewItem(reader("ID").ToString)
                x.SubItems.Add(reader("Room").ToString)
                ListView1.Items.Add(x)
            Loop

        Catch ex As Exception
            MsgBox("error found ;" & ex.Message & ex.StackTrace)
            cnn.Close()
        Finally
            cnn.Close()
        End Try
    End Sub

    Private Sub RoomManagement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Load_Room()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
        Home.switchPanel(RoomAdd)
    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick
        Dim IDRoom As String = ListView1.SelectedItems(0).SubItems(0).Text
        Dim RoomInfo As String = ListView1.SelectedItems(0).SubItems(1).Text

        RoomUpdate.TextBox1.Text = IDRoom
        RoomUpdate.TextBox2.Text = RoomInfo
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Me.Dispose()
        Home.switchPanel(RoomUpdate)
    End Sub
End Class