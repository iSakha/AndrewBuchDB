Module myFunctions
    '===================================================================================
    '             === Clear controls ===
    '===================================================================================
    Sub clearControls()

        mainForm.rtb_fixtureName.Text = ""
        mainForm.txt_qty.Text = ""
        mainForm.rtb_FirstName.Text = ""
        mainForm.txt_qty1.Text = ""
        mainForm.rtb_SecondName.Text = ""
        mainForm.txt_qty2.Text = ""
        mainForm.rtb_ThirdName.Text = ""
        mainForm.txt_qty3.Text = ""

        mainForm.lbl_qty_belimlight.Text = ""
        mainForm.lbl_qty_PRLighting.Text = ""
        mainForm.lbl_qty_blackout.Text = ""
        mainForm.lbl_qty_vision.Text = ""
        mainForm.lbl_qtyTotal.Text = ""
        mainForm.lbl_smeta_qty.Visible = False

    End Sub
    '===================================================================================
    '             === Prev record ===
    '===================================================================================
    Sub prevRecord()

        Dim index As Integer
        Dim selectedRow As DataGridViewRow

        index = mainForm.DGV_light.CurrentRow.Index

        mainForm.DGV_light.ClearSelection()
        mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
        mainForm.DGV_light.Rows(index).Selected = True

        If index = 0 Then
            index = mainForm.DGV_light.Rows.Count
        End If
        index = index - 1
        mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
        mainForm.DGV_light.Rows(index).Selected = True

        Try
            selectedRow = mainForm.DGV_light.Rows(index)

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

        calcQuantity()

    End Sub
    '===================================================================================
    '             === Next record ===
    '===================================================================================
    Sub nextRecord()
        Dim index As Integer
        Dim selectedRow As DataGridViewRow

        index = mainForm.DGV_light.CurrentRow.Index

        mainForm.DGV_light.ClearSelection()
        mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
        mainForm.DGV_light.Rows(index).Selected = True

        If index = mainForm.DGV_light.Rows.Count - 1 Then
            index = -1
        End If
        index = index + 1
        mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
        mainForm.DGV_light.Rows(index).Selected = True
        Try
            selectedRow = mainForm.DGV_light.Rows(index)

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

        calcQuantity()
    End Sub

    Sub calcQuantity()
        Dim index As Integer
        Dim i, j, qty, sum As Integer

        i = mainForm.cmb_category.SelectedIndex

        index = mainForm.DGV_light.CurrentRow.Index

        Dim row As DataRow
        For j = 0 To 3
            sum = 0
            row = mainForm.dt_Lighting(i, j).rows(index)

            If row.Item(4).ToString = "" Then
                qty = 0
            Else qty = CInt(row.Item(4))

            End If

            sum = sum + qty

            If row.Item(6).ToString = "" Then
                qty = 0
            Else qty = CInt(row.Item(6))

            End If

            sum = sum + qty

            If row.Item(8).ToString = "" Then
                qty = 0
            Else qty = CInt(row.Item(8))

            End If

            sum = sum + qty

            mainForm.lblSumQty(j).Text = sum
        Next j

        mainForm.lbl_qtyTotal.Text = mainForm.txt_qty.Text
        mainForm.lbl_smeta_qty.Visible = True

    End Sub

End Module
