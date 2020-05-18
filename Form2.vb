Public Class sumForm
    Private Sub btn_close_Click(sender As Object, e As EventArgs) Handles btn_close.Click
        Me.Close()
    End Sub

    Private Sub sumForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim index As Integer
        Try
            index = mainForm.DGV_light.CurrentRow.Index

            Me.dgv_sum.CurrentCell = Me.dgv_sum.Item(0, index)
            Me.dgv_sum.Rows(index).Selected = True
        Catch
        End Try
    End Sub
End Class