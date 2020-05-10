Module myFunctions

    Sub clearControls()

        mainForm.rtb_fixtureName.Text = ""
        mainForm.txt_qty.Text = ""
        mainForm.rtb_FirstName.Text = ""
        mainForm.txt_qty1.Text = ""
        mainForm.rtb_SecondName.Text = ""
        mainForm.txt_qty2.Text = ""
        mainForm.rtb_ThirdName.Text = ""
        mainForm.txt_qty3.Text = ""

    End Sub

    Sub prevRecord()

        Dim index As Integer
        Dim selectedRow As DataGridViewRow

        index = mainForm.DGV.CurrentRow.Index

        mainForm.DGV.ClearSelection()
        mainForm.DGV.CurrentCell = mainForm.DGV.Item(0, index)
        mainForm.DGV.Rows(index).Selected = True

        If index = 0 Then
            index = mainForm.DGV.Rows.Count
        End If
        index = index - 1
        mainForm.DGV.CurrentCell = mainForm.DGV.Item(0, index)
        mainForm.DGV.Rows(index).Selected = True

        Try
            selectedRow = mainForm.DGV.Rows(index)

            mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
            mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
            mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
            mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
            mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
            mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
            mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
            mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString
        Catch

        End Try

    End Sub

    Sub nextRecord()
        Dim index As Integer
        Dim selectedRow As DataGridViewRow

        index = mainForm.DGV.CurrentRow.Index

        mainForm.DGV.ClearSelection()
        mainForm.DGV.CurrentCell = mainForm.DGV.Item(0, index)
        mainForm.DGV.Rows(index).Selected = True

        If index = mainForm.DGV.Rows.Count - 1 Then
            index = -1
        End If
        index = index + 1
        mainForm.DGV.CurrentCell = mainForm.DGV.Item(0, index)
        mainForm.DGV.Rows(index).Selected = True
        Try
            selectedRow = mainForm.DGV.Rows(index)

            mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
            mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
            mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
            mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
            mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
            mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
            mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
            mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString
        Catch

        End Try
    End Sub


End Module
