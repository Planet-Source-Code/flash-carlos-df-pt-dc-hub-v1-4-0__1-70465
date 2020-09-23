Public Class frmMain

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If g_Enabled Then _
             chkEnabledPlg.Checked = True _
        Else chkEnabledPlg.Checked = False
    End Sub

    Private Sub chkEnabledPlg_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEnabledPlg.CheckedChanged
        If chkEnabledPlg.Checked = True Then _
             g_Main.RunEvent("Switch", True) _
        Else g_Main.RunEvent("Switch", False)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox(frmHub.RunFunction("HubUpTime"))
    End Sub

End Class