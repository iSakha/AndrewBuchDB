﻿Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Module myFunctions
    '===================================================================================
    '             === Load database ===
    '===================================================================================
    Sub loadDataBaseFolder()

        '   1.   Get database folder over Folder browser

        mainForm.FBD.SelectedPath = Directory.GetCurrentDirectory()
        If (mainForm.FBD.ShowDialog() = DialogResult.OK) Then
            mainForm.sDir_DB = mainForm.FBD.SelectedPath
        End If

        '   2.   Get list of database files , names of each file, list of worksheets in each file

        Dim cat As String                           ' variable to get name of database file

        Dim i As Integer = 1

        mainForm.dbFiles = New Collection           ' collection of database files in db folder
        mainForm.fileNames = New Collection         ' collection of names of each file

        mainForm.mainDict = New Dictionary(Of String, Collection)

        For Each foundFile As String In My.Computer.FileSystem.GetFiles _
            (mainForm.sDir_DB, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.xlsx")

            '   Extract file name from full path

            mainForm.sFileName_DB = CStr(foundFile)         ' full path to database file
            Dim SplitFileName_DB() As String
            SplitFileName_DB = Split(mainForm.sFileName_DB, "\")
            cat = SplitFileName_DB(SplitFileName_DB.Count - 1)
            SplitFileName_DB = Split(cat, ".")
            cat = SplitFileName_DB(0)

            mainForm.fileNames.Add(cat)             ' add name of each file to name collection

            '   Create collection of Excel files

            Dim excelFile = New FileInfo(mainForm.sFileName_DB)

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

            mainForm.dbFiles.Add(excelFile)

            '   Create collection of Excel sheets in each file (i.e. categories)

            Dim j As Integer = 0
            Dim ws As ExcelWorksheet

            mainForm.wsCategory = New Collection

            For j = 0 To Excel.Workbook.Worksheets.Count - 1
                ws = Excel.Workbook.Worksheets(j)
                mainForm.wsCategory.Add(ws)

            Next j

            i = i + 1

            ' Create main dictionary where key value is name of file,
            ' and value is collection of worksheets
            mainForm.mainDict.Add(cat, mainForm.wsCategory)
        Next

        '------------------------------------------------------------------------
        ' Now we can test it to print all files,worksheets and excel tables
        '------------------------------------------------------------------------

        i = 0
        For Each foundFile As String In My.Computer.FileSystem.GetFiles _
            (mainForm.sDir_DB, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.xlsx")

            'Console.WriteLine(mainForm.fileNames(i + 1))
            Dim category As String
            Dim k As Integer = 0
            For Each ws As ExcelWorksheet In mainForm.wsCategory
                category = mainForm.mainDict.Item(mainForm.fileNames(i + 1)).Item(k + 1).Name
                'Console.WriteLine(vbTab & category)
                k = k + 1
            Next ws
            i = i + 1
        Next foundFile
    End Sub
    '===================================================================================
    '             === Get dictionary of all excel tables ===
    '===================================================================================
    Sub load_dbFile(_tabIndex As Integer)

        'If Not (mainForm.sDir_DB = Nothing) Then

        '    ' mainForm.sFileName_DB = mainForm.sDir_DB & _fileName

        '    ' Console.WriteLine(mainForm.sFileName_DB)

        '    Dim excelFile = New FileInfo(mainForm.sFileName_DB)

        '    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        '    Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

        '    mainForm.obj_excel = Excel                            '   Global vars to use in function "Save"
        '    mainForm.obj_excelFile = excelFile
        'End If

        ' 1. Add items (categories) to categoryComboBox

        Dim category As String
        Dim k As Integer = 0
        Dim ws As ExcelWorksheet
        mainForm.cmb_category.Items.Clear()
        For Each ws In mainForm.wsCategory
            category = mainForm.mainDict.Item(mainForm.fileNames(_tabIndex)).Item(k + 1).Name
            'Console.WriteLine(_tabIndex)
            'Console.WriteLine(mainForm.mainDict.Keys(_tabIndex - 1))
            mainForm.cmb_category.Items.Add(category)
            k = k + 1
        Next ws

        ' 2. Get all excel tables from selected file (tab index)

        mainForm.xlTablesDict = New Dictionary(Of String, Collection)
        ' mainForm.xlTables = New Collection

        mainForm.xl_sumTablesDict = New Dictionary(Of String, Collection)


        Dim i, j As Integer

        For i = 1 To mainForm.fileNames.Count
            For j = 1 To mainForm.wsCategory.Count
                ws = mainForm.mainDict.Item(mainForm.fileNames(i)).Item(j)
                mainForm.xlTables = New Collection
                mainForm.xl_sumTables = New Collection
                For Each tbl As ExcelTable In ws.Tables
                    mainForm.xlTables.Add(tbl)
                    'Console.WriteLine(tbl.Name)
                Next tbl
                Dim xl_sumTbl As ExcelTable
                xl_sumTbl = ws.Tables.Item(0)
                Console.WriteLine(xl_sumTbl.Name)
                mainForm.xl_sumTables.Add(xl_sumTbl)
                mainForm.xlTablesDict.Add(ws.Name, mainForm.xlTables)
            Next j
            mainForm.xl_sumTablesDict.Add(ws.Name, mainForm.xl_sumTables)
        Next i

        '------------------------------------------------------------------------
        ' Now we can test it to print all excel tables from xlTablesDict
        '------------------------------------------------------------------------
        For i = 1 To mainForm.fileNames.Count
            'Console.WriteLine(mainForm.fileNames(i))
            For j = 1 To mainForm.wsCategory.Count
                ws = mainForm.mainDict.Item(mainForm.fileNames(i)).Item(j)
                Dim cat As String = mainForm.mainDict.Item(mainForm.fileNames(i)).Item(j).Name
                Dim tblName As String
                'Console.WriteLine(vbTab & cat)
                For Each tbl As ExcelTable In ws.Tables
                    tblName = tbl.Name
                    'Console.WriteLine(vbTab & vbTab & tblName)
                Next tbl
            Next j

            Next i

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
        Dim i As Integer = mainForm.tabControl.SelectedIndex - 1

        index = mainForm.dgv(i).CurrentRow.Index

        mainForm.dgv(i).ClearSelection()
        'sumForm.dgv_sum.ClearSelection()
        mainForm.dgv(i).CurrentCell = mainForm.dgv(i).Item(0, index)
        mainForm.dgv(i).Rows(index).Selected = True


        If index = 0 Then
            index = mainForm.dgv(i).Rows.Count - 1
        End If



        index = index - 1
            mainForm.dgv(i).CurrentCell = mainForm.dgv(i).Item(0, index)
            mainForm.dgv(i).Rows(index).Selected = True
        'sumForm.dgv_sum.Rows(index).Selected = True

        selectedRow = mainForm.dgv(i).Rows(index)

        mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
            mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
            mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
            mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
            mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
            mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
            mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
            mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString


        'calcQuantity()

    End Sub
    '===================================================================================
    '             === Next record ===
    '===================================================================================
    Sub nextRecord()
        Dim index As Integer
        Dim selectedRow As DataGridViewRow
        Dim i As Integer = mainForm.tabControl.SelectedIndex - 1

        index = mainForm.dgv(i).CurrentRow.Index

        mainForm.dgv(i).ClearSelection()
        'sumForm.dgv_sum.ClearSelection()
        mainForm.dgv(i).CurrentCell = mainForm.dgv(i).Item(0, index)
        mainForm.dgv(i).Rows(index).Selected = True

        If index = mainForm.dgv(i).Rows.Count - 2 Then
            index = -1
        End If



        index = index + 1
            mainForm.dgv(i).CurrentCell = mainForm.dgv(i).Item(0, index)
            mainForm.dgv(i).Rows(index).Selected = True
        'sumForm.dgv_sum.Rows(index).Selected = True

        selectedRow = mainForm.dgv(i).Rows(index)

            mainForm.rtb_fixtureName.Text = selectedRow.Cells(1).Value.ToString
            mainForm.txt_qty.Text = selectedRow.Cells(2).Value.ToString
            mainForm.rtb_FirstName.Text = selectedRow.Cells(3).Value.ToString
            mainForm.txt_qty1.Text = selectedRow.Cells(4).Value.ToString
            mainForm.rtb_SecondName.Text = selectedRow.Cells(5).Value.ToString
            mainForm.txt_qty2.Text = selectedRow.Cells(6).Value.ToString
            mainForm.rtb_ThirdName.Text = selectedRow.Cells(7).Value.ToString
            mainForm.txt_qty3.Text = selectedRow.Cells(8).Value.ToString


    End Sub
    '===================================================================================
    '             === Calculate quantity ===
    '===================================================================================
    Sub calcQuantity()

        Dim index As Integer
        Dim i, j, qty, sum As Integer

        i = mainForm.cmb_category.SelectedIndex

        index = mainForm.dgv(mainForm.tabControl.SelectedIndex - 1).CurrentRow.Index

        For j = 0 To mainForm.sCompany.Count - 1
            sum = 0

            qty = mainForm.dtDict.Item(mainForm.tabControl.SelectedIndex).Item(j + 1).Rows(index).Item(4)
            sum = sum + qty

            qty = mainForm.dtDict.Item(mainForm.tabControl.SelectedIndex).Item(j + 1).Rows(index).Item(6)
            sum = sum + qty

            qty = mainForm.dtDict.Item(mainForm.tabControl.SelectedIndex).Item(j + 1).Rows(index).Item(8)
            sum = sum + qty

            mainForm.lblSumQty(j).Text = sum
            mainForm.lblSumQty(j).Visible = True

        Next j

        mainForm.lbl_qtyTotal.Text = mainForm.txt_qty.Text
        mainForm.lbl_smeta_qty.Visible = True

    End Sub

    '===================================================================================
    '             === ADD data to DB ===
    '===================================================================================
    Sub addData(_catIndex As Integer)

        Dim sRow(4, 7) As String
        sRow = New String(4, 7) {}
        Dim rCount As Integer
        Dim dt As DataTable
        Dim row As DataRow

        sRow = {
            {newForm.txt_name_addform.Text, newForm.txt_qty_addform.Text,
            newForm.txt_belimlight1_addform.Text, newForm.txt_qty_belimlight1_addform.Text,
            newForm.txt_belimlight2_addform.Text, newForm.txt_qty_belimlight2_addform.Text,
            newForm.txt_belimlight3_addform.Text, newForm.txt_qty_belimlight3_addform.Text},
            {newForm.txt_name_addform.Text, newForm.txt_qty_addform.Text,
            newForm.txt_PRlighting1_addform.Text, newForm.txt_qty_PRlighting1_addform.Text,
            newForm.txt_PRlighting2_addform.Text, newForm.txt_qty_PRlighting2_addform.Text,
            newForm.txt_PRlighting3_addform.Text, newForm.txt_qty_PRlighting3_addform.Text},
            {newForm.txt_name_addform.Text, newForm.txt_qty_addform.Text,
            newForm.txt_blackout1_addform.Text, newForm.txt_qty_blackout1_addform.Text,
            newForm.txt_blackout2_addform.Text, newForm.txt_qty_blackout2_addform.Text,
            newForm.txt_blackout3_addform.Text, newForm.txt_qty_blackout3_addform.Text},
            {newForm.txt_name_addform.Text, newForm.txt_qty_addform.Text,
            newForm.txt_vision1_addform.Text, newForm.txt_qty_vision1_addform.Text,
            newForm.txt_vision2_addform.Text, newForm.txt_qty_vision2_addform.Text,
            newForm.txt_vision3_addform.Text, newForm.txt_qty_vision3_addform.Text},
            {newForm.txt_name_addform.Text, newForm.txt_qty_addform.Text,
            newForm.txt_stage1_addform.Text, newForm.txt_qty_stage1_addform.Text,
            newForm.txt_stage2_addform.Text, newForm.txt_qty_stage2_addform.Text,
            newForm.txt_stage3_addform.Text, newForm.txt_qty_stage3_addform.Text}
        }


        For j As Integer = 0 To mainForm.sCompany.Count - 1

            'dt = mainForm.dt_Lighting(_catIndex, j)
            'rCount = dt.Rows.Count
            'row = dt.Rows.Add()

            row.Item(0) = CInt(dt.Rows(rCount - 1).Item(0)) + 1
            row.Item(1) = sRow(j, 0)
            row.Item(2) = CInt(sRow(j, 1))
            row.Item(3) = sRow(j, 2)
            row.Item(4) = CInt(sRow(j, 3))
            row.Item(5) = sRow(j, 4)
            row.Item(6) = CInt(sRow(j, 5))
            row.Item(7) = sRow(j, 6)
            row.Item(8) = CInt(sRow(j, 7))

        Next j

        rCount = mainForm.dt_sumLighting(_catIndex).Rows.Count
        row = mainForm.dt_sumLighting(_catIndex).Rows.Add()
        'update_sumDatatable(rCount)

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
        '   Chek null values in textboxes

        For i As Integer = 1 To sRow.Count - 1 Step 2
            If sRow(i) = "" Then
                MsgBox("Поле количества приборов не может быть пустым!")
                mainForm.btn_save.Enabled = False
                Exit Sub
            End If
        Next i

        For colIndex As Integer = 1 To 8
            row.Item(colIndex) = sRow(colIndex - 1)
        Next colIndex
        _dgv.DataSource = _dt

        ' update_sumDatatable(_index)

    End Sub

    '===================================================================================
    '             === DELETE data from DB ===
    '===================================================================================

    Sub deleteData(_catIndex As Integer, _rowIndex As Integer)
        Dim rowCollection As DataRowCollection
        Dim j As Integer

        For j = 0 To mainForm.sCompany.Count - 1
            '        rowCollection = mainForm.dt_Lighting(_catIndex, j).rows
            rowCollection.RemoveAt(_rowIndex)
        Next

        rowCollection = mainForm.dt_sumLighting(_catIndex).rows
        rowCollection.RemoveAt(_rowIndex)

        ' mainForm.DGV_light.DataSource = mainForm.dt_Lighting(_catIndex, mainForm.selCompIndex)
    End Sub

    '===================================================================================
    '             === SAVE data to DB ===
    '===================================================================================

    Sub saveData(_i As Integer, _j As Integer)

        Dim startCell As String
        Dim oldAddr As OfficeOpenXml.ExcelAddressBase
        Dim newAddr As OfficeOpenXml.ExcelAddressBase

        Select Case mainForm.selEditModeIndex

            '           "Update" selected
            Case 0
                startCell = mainForm.tbl_Lighting_tables(_i, _j).Address.Start.Address
                mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_Lighting(_i, _j), True)

            '           "Delete" selected
            Case 1

                Dim j As Integer

                For j = 0 To mainForm.sCompany.Count - 1

                    startCell = mainForm.tbl_Lighting_tables(_i, j).Address.Start.Address



                    'Console.WriteLine(mainForm.tbl_Lighting_tables(_i, j).Range.End.Row)

                    oldAddr = mainForm.tbl_Lighting_tables(_i, j).Address
                    newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row - 1, oldAddr.End.Column)
                    mainForm.tbl_Lighting_tables(_i, j).TableXml.InnerXml = mainForm.tbl_Lighting_tables(_i, j).
                        TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                    mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_Lighting(_i, j), True)
                Next j

                startCell = mainForm.tbl_Lighting_sumTables(_i).Address.Start.Address


                oldAddr = mainForm.tbl_Lighting_sumTables(_i).Address
                newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row - 1, oldAddr.End.Column)
                mainForm.tbl_Lighting_sumTables(_i).TableXml.InnerXml = mainForm.tbl_Lighting_sumTables(_i).
                        TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_sumLighting(_i), True)



            '           "Add" selected

            Case 2
                Dim j As Integer
                For j = 0 To mainForm.sCompany.Count - 1
                    startCell = mainForm.tbl_Lighting_tables(_i, j).Address.Start.Address


                    oldAddr = mainForm.tbl_Lighting_tables(_i, j).Address
                    newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row + 1, oldAddr.End.Column)
                    mainForm.tbl_Lighting_tables(_i, j).TableXml.InnerXml = mainForm.tbl_Lighting_tables(_i, j).
                        TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                    mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_Lighting(_i, j), True)
                Next j

                startCell = mainForm.tbl_Lighting_sumTables(_i).Address.Start.Address


                oldAddr = mainForm.tbl_Lighting_sumTables(_i).Address
                newAddr = New ExcelAddressBase(oldAddr.Start.Row, oldAddr.Start.Column, oldAddr.End.Row + 1, oldAddr.End.Column)
                mainForm.tbl_Lighting_sumTables(_i).TableXml.InnerXml = mainForm.tbl_Lighting_sumTables(_i).
                        TableXml.InnerXml.Replace(oldAddr.ToString(), newAddr.ToString())

                mainForm.wsLight(_i).Cells(startCell).LoadFromDataTable(mainForm.dt_sumLighting(_i), True)

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
