Public Class Home

    Sub switchPanel(ByVal panel As Form)
        FormPanel.Controls.Clear()
        panel.TopLevel = False
        FormPanel.Controls.Add(panel)
        panel.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        switchPanel(ScheduleManagement)
        ScheduleManagement.Load_Schedule()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        switchPanel(RoomManagement)
        RoomManagement.Load_Room()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        switchPanel(SubjectManagement)
        SubjectManagement.Load_Subject()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        switchPanel(FacultyManagement)
        FacultyManagement.Load_Faculty()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        switchPanel(StrandManagement)
        StrandManagement.Load_Section()
    End Sub
End Class
