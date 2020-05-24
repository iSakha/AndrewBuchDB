Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Public Class mainForm

    Public sDir_DB As String                        ' database folder
    Public sFileName_DB As String                   ' full path to database file

    Public wsLight() As ExcelWorksheet
    Public wsScreen() As ExcelWorksheet

    '=================================  ExcelTable  ========================================================

    Public tbl_Lighting_tables(7, 4) As ExcelTable
    Public tbl_Lighting_sumTables(7) As ExcelTable

    Public tbl_Screen_tables(7, 4) As ExcelTable
    Public tbl_Screen_sumTables(7) As ExcelTable



    '=================================  DataTable  ========================================================

    'Public dt_Lighting(7, 4) As Object
    Public dt_sumLighting(7) As Object

    'Public dt_Screen(7, 4) As Object
    Public dt_sumScreen(7) As Object

    Public dtColl As Collection
    Public dtDict As Dictionary(Of Integer, Collection)

    '=================================  Rows and columns  ==================================================

    Public r_Light_tbl(7, 4) As Integer
    Public r_Light_sumTbl(7) As Integer

    Public r_Screen_tbl(7, 4) As Integer
    Public r_Screen_sumTbl(7) As Integer


    Public c_Light_tbl(7, 4) As Integer
    Public c_Light_sumTbl(7) As Integer

    Public c_Screen_tbl(7, 4) As Integer
    Public c_Screen_sumTbl(7) As Integer

    '=================================      Address        ==================================================

    Public adr_Light_tbl(7, 4) As String
    Public adr_Light_sumTbl(7) As String

    Public adr_Screen_tbl(7, 4) As String
    Public adr_Screen_sumTbl(7) As String

    '=================================      ExcelRange       =================================================

    Public rng_Light_tbl(7, 4) As ExcelRange
    Public rng_Light_sumTbl(7) As ExcelRange

    Public rng_Screen_tbl(7, 4) As ExcelRange
    Public rng_Screen_sumTbl(7) As ExcelRange

    '=================================      Others       =================================================

    Public obj_excel, obj_excelFile As Object         '   Global vars to use in function "Save"

    Public sCompany() As String = {"belimlight", "PRLighting", "blackout", "vision", "stage"}
    Public selCompIndex As Integer = 0

    Public lblSumQty() As Object
    Public dgv() As Object

    Public editMode() As String = {"Update", "Delete", "Add"}
    Public selEditModeIndex As Integer = 0
    Public selectedCategoryIndex As Integer

    Public dbFiles, wsCategory, xlTables, fileNames As Collection
    Public mainDict As Dictionary(Of String, Collection)        'key - filename,collection - categories(Excel worksheets)
    Public xlTablesDict As Dictionary(Of String, Collection)    'key - category, collection - xlTables


    '===================================================================================      
    '                === Load button ===
    '===================================================================================
    Private Sub btn_loadDB_Click(sender As Object, e As EventArgs) Handles btn_loadDB.Click

        loadDataBaseFolder()                      '   myFunctions.vb
        initLabels()                              '   declarations.vb
        initDGV()                                 '   declarations.vb
        'dt_Lighting(7, 4) = New Object
        'dt_Screen(7, 4) = New Object
    End Sub
    '===================================================================================
    '             === Select page ===
    '===================================================================================
    Private Sub tabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabControl.SelectedIndexChanged


        If tabControl.SelectedIndex > 0 Then

            load_dbFile(tabControl.SelectedIndex)
            cmb_category.SelectedIndex = 0

        End If

        grbx_1.Visible = True
        grbx_2.Visible = True
    End Sub

    '===================================================================================      
    '                === Belimlight button ===
    '===================================================================================
    Private Sub btn_belIm_Click_1(sender As Object, e As EventArgs) Handles btn_belIm.Click

        selCompIndex = 0
        btn_prev.Enabled = True
        btn_next.Enabled = True
        writeZeroInQtyTxt()
        Dim c As Color = Color.FromArgb(252, 228, 214)

        '           dTable.vb
        create_datatable(cmb_category.Items(cmb_category.SelectedIndex), 1)

        Dim i, j As Integer

        i = cmb_category.SelectedIndex
        j = tabControl.SelectedIndex


        dgv(j - 1).DataSource = dtDict.Item(j).Item(1)
        DGV_format(c)


        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        clearControls()

    End Sub

    '===================================================================================      
    '                === PRLighting button ===
    '===================================================================================
    Private Sub btn_prLight_Click_1(sender As Object, e As EventArgs) Handles btn_prLight.Click

        selCompIndex = 1
        btn_prev.Enabled = True
        btn_next.Enabled = True
        writeZeroInQtyTxt()
        Dim c As Color = Color.FromArgb(221, 235, 247)

        '           dTable.vb
        create_datatable(cmb_category.Items(cmb_category.SelectedIndex), 2)

        Dim i, j As Integer

        i = cmb_category.SelectedIndex
        j = tabControl.SelectedIndex

        dgv(j - 1).DataSource = dtDict.Item(j).Item(1)
        DGV_format(c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        clearControls()

    End Sub

    '===================================================================================      
    '                === Blackout button ===
    '===================================================================================
    Private Sub btn_blackOut_Click_1(sender As Object, e As EventArgs) Handles btn_blackOut.Click

        selCompIndex = 2
        btn_prev.Enabled = True
        btn_next.Enabled = True
        writeZeroInQtyTxt()
        Dim c As Color = Color.FromArgb(237, 237, 237)

        '           dTable.vb
        create_datatable(cmb_category.Items(cmb_category.SelectedIndex), 3)

        Dim i, j As Integer

        i = cmb_category.SelectedIndex
        j = tabControl.SelectedIndex

        dgv(j - 1).DataSource = dtDict.Item(j).Item(1)
        DGV_format(c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        clearControls()

    End Sub

    '===================================================================================      
    '                === Vision button ===  
    '===================================================================================
    Private Sub btn_vision_Click_1(sender As Object, e As EventArgs) Handles btn_vision.Click

        selCompIndex = 3
        btn_prev.Enabled = True
        btn_next.Enabled = True
        writeZeroInQtyTxt()
        Dim c As Color = Color.FromArgb(226, 239, 218)

        '           dTable.vb
        create_datatable(cmb_category.Items(cmb_category.SelectedIndex), 4)

        Dim i, j As Integer

        i = cmb_category.SelectedIndex
        j = tabControl.SelectedIndex

        dgv(j - 1).DataSource = dtDict.Item(j).Item(1)
        DGV_format(c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        clearControls()

    End Sub
    '===================================================================================      
    '                === Stage button ===  
    '===================================================================================
    Private Sub btn_stage_Click(sender As Object, e As EventArgs) Handles btn_stage.Click

        selCompIndex = 4
        btn_prev.Enabled = True
        btn_next.Enabled = True
        writeZeroInQtyTxt()
        Dim c As Color = Color.FromArgb(237, 226, 246)

        '           dTable.vb
        create_datatable(cmb_category.Items(cmb_category.SelectedIndex), 5)

        Dim i, j As Integer

        i = cmb_category.SelectedIndex
        j = tabControl.SelectedIndex

        dgv(j - 1).DataSource = dtDict.Item(j).Item(1)
        DGV_format(c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        clearControls()

    End Sub

    '===================================================================================
    '             === Select category ===
    '===================================================================================

    Private Sub cmb_category_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cmb_category.SelectedIndexChanged
        clearControls()
        dgv(0).DataSource = Nothing
        dgv(1).DataSource = Nothing
        btn_prev.Enabled = False
        btn_next.Enabled = False
    End Sub

    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Private Sub DGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_light.CellClick
        dgv_clickCell(sender, e)
        'calcQuantity()
    End Sub

    Private Sub DGV_screen_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_screen.CellClick
        dgv_clickCell(sender, e)
        'calcQuantity()
    End Sub

    '===================================================================================
    '             === Prev record ===
    '===================================================================================
    Private Sub btn_prev_Click(sender As Object, e As EventArgs) Handles btn_prev.Click
        prevRecord()
    End Sub
    '===================================================================================
    '             === Next record ===
    '===================================================================================
    Private Sub btn_next_Click(sender As Object, e As EventArgs) Handles btn_next.Click
        nextRecord()
    End Sub
    '===================================================================================
    '             === ADD data to DB ===
    '===================================================================================
    Private Sub btn_add_Click(sender As Object, e As EventArgs) Handles btn_add.Click
        btn_save.Enabled = True
        Dim i As Integer
        selEditModeIndex = 2
        i = cmb_category.SelectedIndex
        selectedCategoryIndex = i

        btn_save.FlatStyle = FlatStyle.Flat
        blockCompanyButtons()
        blockEditButtons()

        newForm.Show()
        newForm.cmb_category_addform.SelectedIndex = i

    End Sub
    '===================================================================================
    '             === UPDATE data in DB ===
    '===================================================================================
    Private Sub btn_update_Click(sender As Object, e As EventArgs) Handles btn_update.Click
        btn_save.Enabled = True
        selEditModeIndex = 0

        Dim i, j, index As Integer

        i = cmb_category.SelectedIndex
        j = selCompIndex
        index = DGV_light.CurrentRow.Index

        updateData(dt_Lighting(i, j), DGV_light, index)

        btn_save.FlatStyle = FlatStyle.Flat
        blockCompanyButtons()
        blockEditButtons()

    End Sub
    '===================================================================================
    '             === DELETE data from DB ===
    '===================================================================================
    Private Sub btn_del_Click(sender As Object, e As EventArgs) Handles btn_del.Click
        btn_save.Enabled = True

        selEditModeIndex = 1

        Dim i, j, rowIndex As Integer

        i = cmb_category.SelectedIndex
        j = selCompIndex
        rowIndex = DGV_light.CurrentRow.Index
        deleteData(i, rowIndex)

        btn_save.FlatStyle = FlatStyle.Flat
        blockCompanyButtons()
        blockEditButtons()
    End Sub

    '===================================================================================
    '             === SAVE data to DB ===
    '===================================================================================
    Private Sub btn_save_Click(sender As Object, e As EventArgs) Handles btn_save.Click

        Dim i, j As Integer

        i = cmb_category.SelectedIndex
        j = selCompIndex

        clearTable(i, j)

        saveData(i, j)

        clearControls()

        btn_save.FlatStyle = FlatStyle.Standard
        unblockCompanyButtons()
        unblockEditButtons()

        Dim excelFile = New FileInfo(sFileName_DB)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

        obj_excel = Excel                            '   Global vars to use in function "Save"
        obj_excelFile = excelFile

        initWorksheets(tabControl.SelectedIndex)
        initExcelTables(tabControl.SelectedIndex)
        formatExcelTable(i, j)

        Select Case selCompIndex
            Case 0
                btn_belIm.PerformClick()
            Case 1
                btn_prLight.PerformClick()
            Case 2
                btn_blackOut.PerformClick()
            Case 3
                btn_vision.PerformClick()
        End Select

    End Sub
    '===================================================================================
    '             === Cancel ===
    '===================================================================================
    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click

        unblockCompanyButtons()
        unblockEditButtons()

        btn_save.FlatStyle = FlatStyle.Standard


    End Sub

    '===================================================================================
    '             === Show summary ===
    '===================================================================================
    Private Sub btn_summary_Click(sender As Object, e As EventArgs) Handles btn_summary.Click
        Dim i As Integer
        i = cmb_category.SelectedIndex
        Console.WriteLine(i)
        sumForm.dgv_sum.DataSource = dt_sumLighting(i)
        sumForm.Show()
        format_sumDGV()

    End Sub
    '===================================================================================
    '             === T E S T   B U T T O N ===
    '===================================================================================
    Private Sub btn_tst_Click(sender As Object, e As EventArgs) Handles btn_tst.Click


        '-------------------------------------------------------------------------------------------
        '               Open file from openFileDialog
        '-------------------------------------------------------------------------------------------

        'sDir_DB = Directory.GetCurrentDirectory()

        'OFD.InitialDirectory = sDir_DB
        'OFD.Title = "Select .omdb file"
        'If OFD.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

        '    sFileName_DB = OFD.FileName

        '    Dim excelFile = New FileInfo(sFileName_DB)

        '    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        '    Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

        '    Dim xlTable As ExcelTable
        '    Dim ws As ExcelWorksheet

        'For i As Integer = 0 To Excel.Workbook.Worksheets.Count - 1
        '    ws = Excel.Workbook.Worksheets(i)

        '    For j As Integer = 0 To ws.Tables.Count - 1
        '        xlTable = ws.Tables.Item(j)
        '        Console.WriteLine(xlTable.Name)
        '    Next j

        'Next i
        'End If
        '-------------------------------------------------------------------------------------------
        '               Open folder from Folder browser
        '-------------------------------------------------------------------------------------------
        sDir_DB = "C:\Users\Sakha\OneDrive\Visual Studio 2019\PROJECTS\AndrewBuch\database"

        Dim name As String
        Dim i As Integer = 1
        dbFiles = New Collection
        fileNames = New Collection
        mainDict = New Dictionary(Of String, Collection)
        Dim ws As ExcelWorksheet

        For Each foundFile As String In My.Computer.FileSystem.GetFiles _
            (sDir_DB, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.xlsx")

            sFileName_DB = CStr(foundFile)

            Dim SplitFileName_DB() As String
            SplitFileName_DB = Split(sFileName_DB, "\")
            name = SplitFileName_DB(SplitFileName_DB.Count - 1)
            Console.WriteLine(name)
            SplitFileName_DB = Split(name, ".")
            name = SplitFileName_DB(0)
            Console.WriteLine(name)
            fileNames.Add(name)
            Dim excelFile = New FileInfo(sFileName_DB)

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

            dbFiles.Add(excelFile)

            Console.WriteLine(dbFiles.Item(i))
            Dim j As Integer = 0

            wsCategory = New Collection
            For j = 0 To Excel.Workbook.Worksheets.Count - 1
                ws = Excel.Workbook.Worksheets(j)
                wsCategory.Add(ws)

                'Console.WriteLine(ws)
            Next j

            i = i + 1
            mainDict.Add(name, wsCategory)

        Next
        '  fileNames(2) -  ScreenDB,  Item(3) - third item in wsCategory collection
        '   with key = ScreenDB, i.e. worksheet "Controllers_1" from workbook "ScreenDB.xlsx"
        'Console.WriteLine(mainDict.Item(fileNames(2)).Item(3))

        Dim xlTbl As ExcelTable

        xlTablesDict = New Dictionary(Of String, Collection)


        For i = 1 To fileNames.Count
            For j = 1 To wsCategory.Count

                ws = mainDict.Item(fileNames(i)).Item(j)
                xlTables = New Collection
                For Each tbl As ExcelTable In ws.Tables

                    xlTables.Add(tbl)
                    Console.WriteLine(tbl.Name)
                Next tbl
                xlTablesDict.Add(ws.Name, xlTables)
            Next j

        Next i

        xlTbl = mainDict.Item(fileNames(2)).Item(3).Tables.Item(2)

        'Console.WriteLine(xlTbl.Name)



    End Sub

End Class
