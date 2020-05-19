Imports OfficeOpenXml
Imports OfficeOpenXml.Table

Module dTable

    '===================================================================================      
    '                === Create summary datatable ===
    '===================================================================================
    Sub create_sumDatatable(_i As Integer)

        Dim i, j As Integer
        Dim adr As String
        Dim rCount, cCount As Integer
        Dim dt As DataTable
        Dim row As DataRow
        Dim rng As ExcelRange
        Dim xlTable As ExcelTable


        xlTable = mainForm.tbl_Lighting_sumTables(_i)
        rCount = xlTable.Address.Rows
        cCount = xlTable.Address.Columns
        adr = xlTable.Address.Address
        rng = mainForm.wsLight(_i).Cells(adr)

        dt = New DataTable

        'Adding the Columns
        For i = 0 To cCount - 1
            dt.Columns.Add(rng.Value(0, i))
        Next i
        dt.TableName = xlTable.Name

        dt.Columns(0).DataType = System.Type.GetType("System.Int32")               ' #
        dt.Columns(1).DataType = System.Type.GetType("System.String")              ' Fixture
        dt.Columns(2).DataType = System.Type.GetType("System.Int32")               ' Q-ty
        dt.Columns(3).DataType = System.Type.GetType("System.Int32")               ' BelImlight
        dt.Columns(4).DataType = System.Type.GetType("System.Int32")               ' PRLightigTouring
        dt.Columns(5).DataType = System.Type.GetType("System.Int32")               ' BlackOut
        dt.Columns(6).DataType = System.Type.GetType("System.Int32")               ' Vision
        dt.Columns(7).DataType = System.Type.GetType("System.Int32")               ' Stage

        'Add Rows from Excel table

        For i = 1 To rCount - 1

            row = dt.Rows.Add()

            For j = 0 To cCount - 1

                row.Item(j) = rng.Value(i, j)

            Next j

        Next i

        mainForm.dt_sumLighting(_i) = dt


    End Sub

    Sub create_sumDatatable_v2(_i As Integer)
        Dim i, j As Integer
        Dim adr As String
        Dim rCount, cCount As Integer
        Dim dt0, dt1, dt2, dt3, dt4, dtSum As DataTable
        Dim row As DataRow
        Dim rng As ExcelRange
        Dim xlTable As ExcelTable


        xlTable = mainForm.tbl_Lighting_sumTables(_i)
        rCount = xlTable.Address.Rows
        cCount = xlTable.Address.Columns
        adr = xlTable.Address.Address
        rng = mainForm.wsLight(_i).Cells(adr)

        dt0 = mainForm.dt_Lighting(_i, 0)
        dt1 = mainForm.dt_Lighting(_i, 1)
        dt2 = mainForm.dt_Lighting(_i, 2)
        dt3 = mainForm.dt_Lighting(_i, 3)
        dt4 = mainForm.dt_Lighting(_i, 4)

        dtSum = New DataTable

        'Adding the Columns
        For i = 0 To cCount - 1
            dtSum.Columns.Add(rng.Value(0, i))
        Next i
        dtSum.TableName = xlTable.Name

        dtSum.Columns(0).DataType = System.Type.GetType("System.Int32")               ' #
        dtSum.Columns(1).DataType = System.Type.GetType("System.String")              ' Fixture
        dtSum.Columns(2).DataType = System.Type.GetType("System.Int32")               ' Q-ty
        dtSum.Columns(3).DataType = System.Type.GetType("System.Int32")               ' BelImlight
        dtSum.Columns(4).DataType = System.Type.GetType("System.Int32")               ' PRLightigTouring
        dtSum.Columns(5).DataType = System.Type.GetType("System.Int32")               ' BlackOut
        dtSum.Columns(6).DataType = System.Type.GetType("System.Int32")               ' Vision
        dtSum.Columns(7).DataType = System.Type.GetType("System.Int32")

        'Add Rows from datatables

        For i = 1 To rCount - 1

            row = dtSum.Rows.Add()

            row.Item(0) = dt0.Rows(i - 1).Item(0)
            row.Item(1) = dt0.Rows(i - 1).Item(1)
            row.Item(2) = dt0.Rows(i - 1).Item(2)
            row.Item(3) = dt0.Rows(i - 1).Item(4) + dt0.Rows(i - 1).Item(6) + dt0.Rows(i - 1).Item(8)
            row.Item(4) = dt1.Rows(i - 1).Item(4) + dt1.Rows(i - 1).Item(6) + dt1.Rows(i - 1).Item(8)
            row.Item(5) = dt2.Rows(i - 1).Item(4) + dt2.Rows(i - 1).Item(6) + dt2.Rows(i - 1).Item(8)
            row.Item(6) = dt3.Rows(i - 1).Item(4) + dt3.Rows(i - 1).Item(6) + dt3.Rows(i - 1).Item(8)
            row.Item(7) = dt4.Rows(i - 1).Item(4) + dt4.Rows(i - 1).Item(6) + dt4.Rows(i - 1).Item(8)

        Next i


        mainForm.dt_sumLighting(_i) = dtSum

    End Sub
    '===================================================================================      
    '                === Update sumDatatable after UPDATE ===
    '===================================================================================
    Sub update_sumDatatable(_i As Integer)
        Dim i As Integer
        Dim dt0, dt1, dt2, dt3, dt4, dtSum As DataTable
        Dim row As DataRow

        i = mainForm.cmb_category.SelectedIndex

        dtSum = mainForm.dt_sumLighting(i)

        row = dtSum.Rows(_i)

        dt0 = mainForm.dt_Lighting(i, 0)
        dt1 = mainForm.dt_Lighting(i, 1)
        dt2 = mainForm.dt_Lighting(i, 2)
        dt3 = mainForm.dt_Lighting(i, 3)
        dt4 = mainForm.dt_Lighting(i, 4)

        row.Item(0) = dt0.Rows(_i).Item(0)
        row.Item(1) = dt0.Rows(_i).Item(1)
        row.Item(2) = dt0.Rows(_i).Item(2)
        row.Item(3) = dt0.Rows(_i).Item(4) + dt0.Rows(_i).Item(6) + dt0.Rows(_i).Item(8)
        row.Item(4) = dt1.Rows(_i).Item(4) + dt1.Rows(_i).Item(6) + dt1.Rows(_i).Item(8)
        row.Item(5) = dt2.Rows(_i).Item(4) + dt2.Rows(_i).Item(6) + dt2.Rows(_i).Item(8)
        row.Item(6) = dt3.Rows(_i).Item(4) + dt3.Rows(_i).Item(6) + dt3.Rows(_i).Item(8)
        row.Item(7) = dt4.Rows(_i).Item(4) + dt4.Rows(_i).Item(6) + dt4.Rows(_i).Item(8)

        mainForm.dt_sumLighting(i) = dtSum

    End Sub

    '===================================================================================      
    '                === Create datatable ===
    '===================================================================================

    Sub create_datatable(_i As Integer, _j As Integer)

        Dim i, j As Integer
        Dim adr As String
        Dim rCount, cCount As Integer
        Dim dt As DataTable
        Dim row As DataRow
        Dim rng As ExcelRange
        Dim xlTable As ExcelTable

        xlTable = mainForm.tbl_Lighting_tables(_i, _j)
        rCount = xlTable.Address.Rows
        cCount = xlTable.Address.Columns
        adr = xlTable.Address.Address
        rng = mainForm.wsLight(_i).Cells(adr)

        dt = New DataTable

        'Adding the Columns
        For i = 0 To cCount - 1
            dt.Columns.Add(rng.Value(0, i))
        Next i
        dt.TableName = xlTable.Name

        dt.Columns(0).DataType = System.Type.GetType("System.Int32")
        dt.Columns(1).DataType = System.Type.GetType("System.String")
        dt.Columns(2).DataType = System.Type.GetType("System.Int32")
        dt.Columns(3).DataType = System.Type.GetType("System.String")
        dt.Columns(4).DataType = System.Type.GetType("System.Int32")
        dt.Columns(5).DataType = System.Type.GetType("System.String")
        dt.Columns(6).DataType = System.Type.GetType("System.Int32")
        dt.Columns(7).DataType = System.Type.GetType("System.String")
        dt.Columns(8).DataType = System.Type.GetType("System.Int32")


        'Add Rows from Excel table

        For i = 1 To rCount - 1
            row = dt.Rows.Add()

            For j = 0 To cCount - 1

                If rng.Value(i, j) = Nothing Then
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
                    row.Item(j) = rng.Value(i, j)
                End If

            Next j
        Next i

        mainForm.dt_Lighting(_i, _j) = dt

    End Sub
    '===================================================================================      
    '                === Format DataGridView ===
    '===================================================================================
    Sub DGV_format(_dtName As String, _color As Color)

        mainForm.DGV_light.Columns(0).Width = 40                ' #
        mainForm.DGV_light.Columns(1).Width = 175               ' Fixture
        mainForm.DGV_light.Columns(2).Width = 40                ' Q-ty
        mainForm.DGV_light.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV_light.Columns(3).Width = 220               ' BelImlight_1  (PRLightigTouring, BlackOut, Vision, Stage)
        mainForm.DGV_light.Columns(4).Width = 40                ' Q-ty_1
        mainForm.DGV_light.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV_light.Columns(5).Width = 220               ' BelImlight_2  (PRLightigTouring, BlackOut, Vision, Stage)
        mainForm.DGV_light.Columns(6).Width = 40                ' Q-ty_2
        mainForm.DGV_light.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        mainForm.DGV_light.Columns(7).Width = 180               ' BelImlight_3  (PRLightigTouring, BlackOut, Vision, Stage)
        mainForm.DGV_light.Columns(8).Width = 40                ' Q-ty_3
        mainForm.DGV_light.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For i = 0 To mainForm.DGV_light.Rows.Count - 2

            'mainForm.DGV_in.Rows(i).Cells(1).Value = Date.FromOADate(mainForm.DGV_in.Rows(i).Cells(1).Value)
            mainForm.DGV_light.RowsDefaultCellStyle.BackColor = _color
            mainForm.DGV_light.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)

        Next i

    End Sub
    Sub format_sumDGV()

        Dim col() As Color

        col = {Color.FromArgb(252, 228, 214), Color.FromArgb(221, 235, 247), Color.FromArgb(237, 237, 237),
            Color.FromArgb(226, 239, 218), Color.FromArgb(237, 226, 246)}

        sumForm.dgv_sum.Columns(0).Width = 55                ' #
        sumForm.dgv_sum.Columns(1).Width = 240               ' Fixture
        sumForm.dgv_sum.Columns(2).Width = 65                ' Q-ty
        sumForm.dgv_sum.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(3).Width = 65                ' BelImlight
        sumForm.dgv_sum.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(4).Width = 65                ' PRLightigTouring
        sumForm.dgv_sum.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(5).Width = 65                ' BlackOut
        sumForm.dgv_sum.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(6).Width = 65                ' Vision
        sumForm.dgv_sum.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        sumForm.dgv_sum.Columns(7).Width = 65                ' Stage
        sumForm.dgv_sum.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


        sumForm.dgv_sum.Columns(3).DefaultCellStyle.BackColor = col(0)
        sumForm.dgv_sum.Columns(4).DefaultCellStyle.BackColor = col(1)
        sumForm.dgv_sum.Columns(5).DefaultCellStyle.BackColor = col(2)
        sumForm.dgv_sum.Columns(6).DefaultCellStyle.BackColor = col(3)
        sumForm.dgv_sum.Columns(7).DefaultCellStyle.BackColor = col(4)

        For i = 0 To sumForm.dgv_sum.Rows.Count - 2

            sumForm.dgv_sum.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)

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

        Try

            mainForm.DGV_light.Rows(index).Selected = True
            sumForm.dgv_sum.ClearSelection()
            sumForm.dgv_sum.Rows(index).Selected = True

            If index = mainForm.DGV_light.Rows.Count - 1 Then
                writeZeroInQtyTxt()
            End If
        Catch

        End Try
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




    Sub changeSumDataTable(_selectedCategoryIndex As Integer)


        Dim i As Integer = _selectedCategoryIndex
        'add data to summary table
        Dim row As DataRow
        Dim iRow() As Integer
        Dim sRow As String
        'Index = mainForm.selectedCategoryIndex
        Dim dt As DataTable
        iRow = New Integer() {newForm.qty_belimlight, newForm.qty_PRlighting,
            newForm.qty_blackout, newForm.qty_vision, newForm.qty_stage}

        dt = mainForm.dt_sumLighting(i)

        row = dt.Rows.Add()

        For k As Integer = 0 To iRow.Count - 1
            row.Item(k + 2) = iRow(k)
        Next k

        Dim xTable As ExcelTable
        xTable = mainForm.tbl_Lighting_sumTables(i)
        Dim startCell As String = xTable.Address.Start.Address

        Dim oldAddr As OfficeOpenXml.ExcelAddressBase
        Dim newAddr As OfficeOpenXml.ExcelAddressBase

        oldAddr = xTable.Address
        newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row + 1, oldAddr.End.Column)
        xTable.TableXml.InnerXml = xTable.TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

        mainForm.wsLight(i).Cells(startCell).LoadFromDataTable(dt, True)

    End Sub

End Module
