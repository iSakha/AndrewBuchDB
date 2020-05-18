Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Module myFunctions
    '===================================================================================
    '             === Load database ===
    '===================================================================================
    Sub loadDataBase()
        mainForm.sDir_DB = Directory.GetCurrentDirectory()

        mainForm.OFD.InitialDirectory = mainForm.sDir_DB
        mainForm.OFD.Title = "Select .omdb file"

        If mainForm.OFD.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            mainForm.sFileName_DB = mainForm.OFD.FileName

            Dim excelFile = New FileInfo(mainForm.sFileName_DB)

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

            mainForm.obj_excel = Excel                            '   Global vars to use in function "Save"
            mainForm.obj_excelFile = excelFile

        End If
    End Sub




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
        mainForm.lbl_qty_stage.Text = ""
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
        sumForm.dgv_sum.ClearSelection()
        mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
        mainForm.DGV_light.Rows(index).Selected = True


        If index = 0 Then
            index = mainForm.DGV_light.Rows.Count - 1
        End If

        Try

            index = index - 1
            mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
            mainForm.DGV_light.Rows(index).Selected = True
            sumForm.dgv_sum.Rows(index).Selected = True

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
        sumForm.dgv_sum.ClearSelection()
        mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
        mainForm.DGV_light.Rows(index).Selected = True

        If index = mainForm.DGV_light.Rows.Count - 2 Then
            index = -1
        End If

        Try

            index = index + 1
            mainForm.DGV_light.CurrentCell = mainForm.DGV_light.Item(0, index)
            mainForm.DGV_light.Rows(index).Selected = True
            sumForm.dgv_sum.Rows(index).Selected = True

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
    '             === Calculate quantity ===
    '===================================================================================
    Sub calcQuantity()

        Dim index As Integer
        Dim i, j, qty, sum As Integer

        i = mainForm.cmb_category.SelectedIndex

        index = mainForm.DGV_light.CurrentRow.Index

        Try
            For j = 0 To mainForm.sCompany.Count - 1
                sum = 0
                qty = mainForm.dt_Lighting(i, j).Rows(index).Item(4)
                sum = sum + qty
                qty = mainForm.dt_Lighting(i, j).Rows(index).Item(6)
                sum = sum + qty
                qty = mainForm.dt_Lighting(i, j).Rows(index).Item(8)
                sum = sum + qty

                mainForm.lblSumQty(j).Text = sum

            Next j
        Catch
        End Try
        mainForm.lbl_qtyTotal.Text = mainForm.txt_qty.Text
        mainForm.lbl_smeta_qty.Visible = True

    End Sub

    '===================================================================================
    '             === ADD data to DB ===
    '===================================================================================
    Sub addData(_dt As DataTable, _dgv As DataGridView)

        Dim rCount As Integer
        Dim sRow() As String
        rCount = _dt.Rows.Count

        sRow = New String() {
                mainForm.rtb_fixtureName.Text,
                mainForm.txt_qty.Text,
                mainForm.rtb_FirstName.Text,
                mainForm.txt_qty1.Text,
                mainForm.rtb_SecondName.Text,
                mainForm.txt_qty2.Text,
                mainForm.rtb_ThirdName.Text,
                mainForm.txt_qty3.Text
                }


        Dim row As DataRow

        row = _dt.Rows.Add()

        For i As Integer = 0 To sRow.Count - 1
            row.Item(i + 1) = sRow(i)
        Next i

        row.Item(0) = CInt(_dt.Rows(rCount - 1).Item(0)) + 1

        _dgv.DataSource = _dt


    End Sub
    '===================================================================================
    '             === UPDATE data in DB ===
    '===================================================================================
    Sub updateData(_dt As DataTable, _dgv As DataGridView, _index As Integer)

        Dim row As DataRow
        row = _dt.Rows(_index)
        Dim sRow() As String



        sRow = New String() {
                mainForm.rtb_fixtureName.Text,
                mainForm.txt_qty.Text,
                mainForm.rtb_FirstName.Text,
                mainForm.txt_qty1.Text,
                mainForm.rtb_SecondName.Text,
                mainForm.txt_qty2.Text,
                mainForm.rtb_ThirdName.Text,
                mainForm.txt_qty3.Text
            }

        For colIndex As Integer = 1 To 8
            row.Item(colIndex) = sRow(colIndex - 1)
        Next colIndex
        _dgv.DataSource = _dt
    End Sub

    '===================================================================================
    '             === DELETE data from DB ===
    '===================================================================================

    Sub deleteData(_dt As DataTable, _dgv As DataGridView, _index As Integer)
        Dim rowCollection As DataRowCollection = _dt.Rows
        rowCollection.RemoveAt(_index)
        _dgv.DataSource = _dt
    End Sub

    '===================================================================================
    '             === SAVE data to DB ===
    '===================================================================================

    Sub saveData(_i As Integer, _j As Integer)

        Select Case mainForm.selEditModeIndex

            '           "Update" selected
            Case 0
                Dim startCell As String = mainForm.tbl_Lighting_tables(_i, _j).Address.Start.Address
                mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_Lighting(_i, _j), True)

            '           "Delete" selected
            Case 1

                Dim j As Integer

                For j = 0 To mainForm.sCompany.Count - 1

                    Dim startCell As String = mainForm.tbl_Lighting_tables(_i, j).Address.Start.Address
                    Dim oldAddr As OfficeOpenXml.ExcelAddressBase
                    Dim newAddr As OfficeOpenXml.ExcelAddressBase


                    Console.WriteLine(mainForm.tbl_Lighting_tables(_i, j).Range.End.Row)

                    oldAddr = mainForm.tbl_Lighting_tables(_i, j).Address
                    newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row - 1, oldAddr.End.Column)
                    mainForm.tbl_Lighting_tables(_i, j).TableXml.InnerXml = mainForm.tbl_Lighting_tables(_i, j).
                        TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                    mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_Lighting(_i, j), True)
                Next j

            '           "Add" selected

            Case 2
                Dim j As Integer
                For j = 0 To mainForm.sCompany.Count - 1
                    Dim startCell As String = mainForm.tbl_Lighting_tables(_i, j).Address.Start.Address
                    Dim oldAddr As OfficeOpenXml.ExcelAddressBase
                    Dim newAddr As OfficeOpenXml.ExcelAddressBase

                    oldAddr = mainForm.tbl_Lighting_tables(_i, j).Address
                    newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row + 1, oldAddr.End.Column)
                    mainForm.tbl_Lighting_tables(_i, j).TableXml.InnerXml = mainForm.tbl_Lighting_tables(_i, j).
                        TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                    mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_Lighting(_i, j), True)
                Next j
        End Select

        mainForm.obj_excel.SaveAs(mainForm.obj_excelFile)

    End Sub

    Sub clearTable(_i As Integer, _j As Integer)

        Dim table As ExcelTable
        Dim firstRow, firstColumn, lastRow, lastColumn As Integer

        table = mainForm.tbl_Lighting_tables(_i, _j)
        firstRow = table.Range.Start.Row
        lastRow = table.Range.End.Row
        firstColumn = table.Range.Start.Column
        lastColumn = table.Range.End.Column

        Dim i, j As Integer
        For i = firstRow + 1 To lastRow

            For j = firstColumn To lastColumn

                mainForm.wsLight(_i).Cells(i, j).Clear()

            Next j

        Next i


    End Sub

    Sub blockCompanyButtons()
        mainForm.btn_belIm.Enabled = False
        mainForm.btn_prLight.Enabled = False
        mainForm.btn_blackOut.Enabled = False
        mainForm.btn_vision.Enabled = False
    End Sub

    Sub unblockCompanyButtons()
        mainForm.btn_belIm.Enabled = True
        mainForm.btn_prLight.Enabled = True
        mainForm.btn_blackOut.Enabled = True
        mainForm.btn_vision.Enabled = True
    End Sub

    Sub blockEditButtons()
        mainForm.btn_add.Enabled = False
        mainForm.btn_update.Enabled = False
        mainForm.btn_del.Enabled = False

    End Sub

    Sub unblockEditButtons()
        mainForm.btn_add.Enabled = True
        mainForm.btn_update.Enabled = True
        mainForm.btn_del.Enabled = True
    End Sub

    Sub writeZeroInQtyTxt()
        If mainForm.txt_qty.Text = "" Then
            mainForm.txt_qty.Text = 0
        End If
        If mainForm.txt_qty1.Text = "" Then
            mainForm.txt_qty1.Text = 0
        End If
        If mainForm.txt_qty2.Text = "" Then
            mainForm.txt_qty2.Text = 0
        End If
        If mainForm.txt_qty3.Text = "" Then
            mainForm.txt_qty3.Text = 0
        End If
    End Sub


    Function isEven(_var As Integer)

        Dim result As Boolean

        If _var Mod 2 = 0 Then
            result = True
        Else
            result = False
        End If
        Return (result)

    End Function
End Module
