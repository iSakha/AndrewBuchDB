Module dTable
    '===================================================================================      
    '                === Create datatable ===
    '===================================================================================

    Sub create_datatable(_rCount As Integer, _colCount As Integer, _rng As Object, _dt As DataTable, _dtName As String)

        Dim i, j As Integer
        Dim row As DataRow



        'Adding the Columns
        For i = 0 To _colCount - 1
            _dt.Columns.Add(_rng.Value(0, i))
        Next i
        _dt.TableName = _dtName

        _dt.Columns(0).DataType = System.Type.GetType("System.String")
        _dt.Columns(1).DataType = System.Type.GetType("System.String")
        _dt.Columns(2).DataType = System.Type.GetType("System.String")
        _dt.Columns(3).DataType = System.Type.GetType("System.String")
        _dt.Columns(4).DataType = System.Type.GetType("System.String")
        _dt.Columns(5).DataType = System.Type.GetType("System.String")
        _dt.Columns(6).DataType = System.Type.GetType("System.String")
        _dt.Columns(7).DataType = System.Type.GetType("System.String")
        _dt.Columns(8).DataType = System.Type.GetType("System.String")

        'Add Rows from Excel table

        For i = 1 To _rCount - 1
            row = _dt.Rows.Add()
            For j = 0 To _colCount - 1
                row.Item(j) = _rng.Value(i, j)
            Next j
        Next i
    End Sub
    '===================================================================================      
    '                === Format DataGridView ===
    '===================================================================================
    Sub DGV_format(_dtName As String, _color As Color)

        mainForm.DGV.Columns(0).Width = 40
        mainForm.DGV.Columns(1).Width = 175
        mainForm.DGV.Columns(2).Width = 40
        mainForm.DGV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(3).Width = 220
        mainForm.DGV.Columns(4).Width = 40
        mainForm.DGV.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(5).Width = 220
        mainForm.DGV.Columns(6).Width = 40
        mainForm.DGV.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV.Columns(7).Width = 180
        mainForm.DGV.Columns(8).Width = 40
        mainForm.DGV.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For i = 0 To mainForm.DGV.Rows.Count - 2

            'mainForm.DGV_in.Rows(i).Cells(1).Value = Date.FromOADate(mainForm.DGV_in.Rows(i).Cells(1).Value)
            mainForm.DGV.RowsDefaultCellStyle.BackColor = _color
            mainForm.DGV.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)

        Next i

    End Sub

    '===================================================================================      
    '                === CellClick on DGV ===
    '===================================================================================
    Sub dgv_clickCell(_sender As Object, _e As DataGridViewCellEventArgs)
        Dim index As Integer
        index = _e.RowIndex
        'Console.WriteLine(index)
        Dim selectedRow As DataGridViewRow
        selectedRow = _sender.Rows(index)

        mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
        mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
        mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
        mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
        mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
        mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
        mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
        mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString

        mainForm.DGV.Rows(index).Selected = True

    End Sub

End Module
