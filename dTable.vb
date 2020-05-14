Imports OfficeOpenXml

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

        _dt.Columns(0).DataType = System.Type.GetType("System.Int32")
        _dt.Columns(1).DataType = System.Type.GetType("System.String")
        _dt.Columns(2).DataType = System.Type.GetType("System.Int32")
        _dt.Columns(3).DataType = System.Type.GetType("System.String")
        _dt.Columns(4).DataType = System.Type.GetType("System.Int32")
        _dt.Columns(5).DataType = System.Type.GetType("System.String")
        _dt.Columns(6).DataType = System.Type.GetType("System.Int32")
        _dt.Columns(7).DataType = System.Type.GetType("System.String")
        _dt.Columns(8).DataType = System.Type.GetType("System.Int32")


        'Add Rows from Excel table

        For i = 1 To _rCount - 1
            row = _dt.Rows.Add()

            For j = 0 To _colCount - 1

                If _rng.Value(i, j) = Nothing Then
                    Select Case j
                        Case 3
                            row.Item(j) = ""
                        Case 4
                            row.Item(j) = 0
                        Case 5
                            row.Item(j) = ""
                        Case 6
                            row.Item(j) = 0
                        Case 7
                            row.Item(j) = ""
                        Case 8
                            row.Item(j) = 0
                    End Select
                Else
                    row.Item(j) = _rng.Value(i, j)
                End If

            Next j
        Next i
    End Sub
    '===================================================================================      
    '                === Format DataGridView ===
    '===================================================================================
    Sub DGV_format(_dtName As String, _color As Color)

        mainForm.DGV_light.Columns(0).Width = 40
        mainForm.DGV_light.Columns(1).Width = 175
        mainForm.DGV_light.Columns(2).Width = 40
        mainForm.DGV_light.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV_light.Columns(3).Width = 220
        mainForm.DGV_light.Columns(4).Width = 40
        mainForm.DGV_light.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV_light.Columns(5).Width = 220
        mainForm.DGV_light.Columns(6).Width = 40
        mainForm.DGV_light.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV_light.Columns(7).Width = 180
        mainForm.DGV_light.Columns(8).Width = 40
        mainForm.DGV_light.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For i = 0 To mainForm.DGV_light.Rows.Count - 2

            'mainForm.DGV_in.Rows(i).Cells(1).Value = Date.FromOADate(mainForm.DGV_in.Rows(i).Cells(1).Value)
            mainForm.DGV_light.RowsDefaultCellStyle.BackColor = _color
            mainForm.DGV_light.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)

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

        mainForm.DGV_light.Rows(index).Selected = True

        If index = mainForm.DGV_light.Rows.Count - 1 Then
            writeZeroInQtyTxt()
        End If

    End Sub


    Sub formatExcelTable(_i As Integer, _j As Integer)

        Dim rng As ExcelRange
        Dim startRow, startColumn, endRow As Integer
        startRow = mainForm.tbl_Lighting_tables(_i, _j).Address.Start.Row
        startColumn = mainForm.tbl_Lighting_tables(_i, _j).Address.Start.Column
        endRow = mainForm.tbl_Lighting_tables(_i, _j).Address.End.Row
        rng = mainForm.wsLight(_i).Cells(startRow + 1, startColumn + 2, endRow, startColumn + 2)
        rng.Style.Numberformat.Format = "0"
        'rng.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
        'rng.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FF0000"))

        rng = mainForm.wsLight(_i).Cells(startRow + 1, startColumn + 4, endRow, startColumn + 4)
        rng.Style.Numberformat.Format = "0"
        'rng.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
        'rng.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FF0000"))

        rng = mainForm.wsLight(_i).Cells(startRow + 1, startColumn + 6, endRow, startColumn + 6)
        rng.Style.Numberformat.Format = "0"
        'rng.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
        'rng.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FF0000"))

        rng = mainForm.wsLight(_i).Cells(startRow + 1, startColumn + 8, endRow, startColumn + 8)
        rng.Style.Numberformat.Format = "0"
        'rng.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
        'rng.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FF0000"))

        mainForm.obj_excel.SaveAs(mainForm.obj_excelFile)

    End Sub

End Module
