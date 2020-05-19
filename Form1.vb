Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Public Class mainForm

    Public sDir_DB As String
    Public sFileName_DB As String

    Public wsMovHeads As ExcelWorksheet
    Public wsStrobes As ExcelWorksheet
    Public wsBlinders As ExcelWorksheet
    Public wsArch As ExcelWorksheet
    Public wsLED As ExcelWorksheet
    Public wsSmoke As ExcelWorksheet
    Public wsConsoles As ExcelWorksheet
    Public wsIntercom As ExcelWorksheet

    Public wsLight() As ExcelWorksheet

    Public tbl_Light_Collection As ExcelTableCollection

    '=================================  ExcelTable  ========================================================

    Public tbl_movHeads_belimlight, tbl_movHeads_PRLighting As ExcelTable
    Public tbl_movHeads_blackout, tbl_movHeads_vision, tbl_movHeads_stage As ExcelTable
    Public tbl_movHeads As ExcelTable

    Public tbl_strobes_belimlight, tbl_strobes_PRLighting As ExcelTable
    Public tbl_strobes_blackout, tbl_strobes_vision, tbl_strobes_stage As ExcelTable
    Public tbl_strobes As ExcelTable

    Public tbl_blinders_belimlight, tbl_blinders_PRLighting As ExcelTable
    Public tbl_blinders_blackout, tbl_blinders_vision, tbl_blinders_stage As ExcelTable
    Public tbl_blinders As ExcelTable

    Public tbl_arch_belimlight, tbl_arch_PRLighting As ExcelTable
    Public tbl_arch_blackout, tbl_arch_vision, tbl_arch_stage As ExcelTable
    Public tbl_arch As ExcelTable

    Public tbl_LED_belimlight, tbl_LED_PRLighting As ExcelTable
    Public tbl_LED_blackout, tbl_LED_vision, tbl_LED_stage As ExcelTable
    Public tbl_LED As ExcelTable

    Public tbl_smoke_belimlight, tbl_smoke_PRLighting As ExcelTable
    Public tbl_smoke_blackout, tbl_smoke_vision, tbl_smoke_stage As ExcelTable
    Public tbl_smoke As ExcelTable

    Public tbl_consoles_belimlight, tbl_consoles_PRLighting As ExcelTable
    Public tbl_consoles_blackout, tbl_consoles_vision, tbl_consoles_stage As ExcelTable
    Public tbl_consoles As ExcelTable

    Public tbl_intercom_belimlight, tbl_intercom_PRLighting As ExcelTable
    Public tbl_intercom_blackout, tbl_intercom_vision, tbl_intercom_stage As ExcelTable
    Public tbl_intercom As ExcelTable



    Public tbl_Lighting_tables(7, 4) As ExcelTable
    Public tbl_Lighting_sumTables(7) As ExcelTable

    '=================================  DataTable  ========================================================

    Public dt_movHeads_belimlight, dt_movHeads_PRLighting As New DataTable()
    Public dt_movHeads_blackout, dt_movHeads_vision, dt_movHeads_stage As New DataTable()
    Public dt_movHeads As New DataTable()

    Public dt_strobes_belimlight, dt_strobes_PRLighting As New DataTable()
    Public dt_strobes_blackout, dt_strobes_vision, dt_strobes_stage As New DataTable()
    Public dt_strobes As New DataTable()

    Public dt_blinders_belimlight, dt_blinders_PRLighting As New DataTable()
    Public dt_blinders_blackout, dt_blinders_vision, dt_blinders_stage As New DataTable()
    Public dt_blinders As New DataTable()

    Public dt_arch_belimlight, dt_arch_PRLighting As New DataTable()
    Public dt_arch_blackout, dt_arch_vision, dt_arch_stage As New DataTable()
    Public dt_arch As New DataTable()

    Public dt_LED_belimlight, dt_LED_PRLighting As New DataTable()
    Public dt_LED_blackout, dt_LED_vision, dt_LED_stage As New DataTable()
    Public dt_LED As New DataTable()

    Public dt_smoke_belimlight, dt_smoke_PRLighting As New DataTable()
    Public dt_smoke_blackout, dt_smoke_vision, dt_smoke_stage As New DataTable()
    Public dt_smoke As New DataTable()

    Public dt_consoles_belimlight, dt_consoles_PRLighting As New DataTable()
    Public dt_consoles_blackout, dt_consoles_vision, dt_consoles_stage As New DataTable()
    Public dt_consoles As New DataTable()

    Public dt_intercom_belimlight, dt_intercom_PRLighting As New DataTable()
    Public dt_intercom_blackout, dt_intercom_vision, dt_intercom_stage As New DataTable()
    Public dt_intercom As New DataTable()

    Public dt_Lighting(7, 4) As Object
    Public dt_sumLighting(7) As Object

    '=================================  Rows and columns  ==================================================

    Public r_movHeads_belimlight, r_movHeads_PRLighting As Integer
    Public r_movHeads_blackout, r_movHeads_vision, r_movHeads_stage As Integer
    Public r_movHeads As Integer

    Public r_strobes_belimlight, r_strobes_PRLighting As Integer
    Public r_strobes_blackout, r_strobes_vision, r_strobes_stage As Integer
    Public r_strobes As Integer

    Public r_blinders_belimlight, r_blinders_PRLighting As Integer
    Public r_blinders_blackout, r_blinders_vision, r_blinders_stage As Integer
    Public r_blinders As Integer

    Public r_arch_belimlight, r_arch_PRLighting As Integer
    Public r_arch_blackout, r_arch_vision, r_arch_stage As Integer
    Public r_arch As Integer

    Public r_LED_belimlight, r_LED_PRLighting As Integer
    Public r_LED_blackout, r_LED_vision, r_LED_stage As Integer
    Public r_LED As Integer

    Public r_smoke_belimlight, r_smoke_PRLighting As Integer
    Public r_smoke_blackout, r_smoke_vision, r_smoke_stage As Integer
    Public r_smoke As Integer

    Public r_consoles_belimlight, r_consoles_PRLighting As Integer
    Public r_consoles_blackout, r_consoles_vision, r_consoles_stage As Integer
    Public r_consoles As Integer

    Public r_intercom_belimlight, r_intercom_PRLighting As Integer
    Public r_intercom_blackout, r_intercom_vision, r_intercom_stage As Integer
    Public r_intercom As Integer

    Public r_Light_tbl(7, 4) As Integer
    Public r_Light_sumTbl(7) As Integer


    Public c_movHeads_belimlight, c_movHeads_PRLighting As Integer
    Public c_movHeads_blackout, c_movHeads_vision, c_movHeads_stage As Integer
    Public c_movHeads As Integer

    Public c_strobes_belimlight, c_strobes_PRLighting As Integer
    Public c_strobes_blackout, c_strobes_vision, c_strobes_stage As Integer
    Public c_strobes As Integer

    Public c_blinders_belimlight, c_blinders_PRLighting As Integer
    Public c_blinders_blackout, c_blinders_vision, c_blinders_stage As Integer
    Public c_blinders As Integer

    Public c_arch_belimlight, c_arch_PRLighting As Integer
    Public c_arch_blackout, c_arch_vision, c_arch_stage As Integer
    Public c_arch As Integer

    Public c_LED_belimlight, c_LED_PRLighting As Integer
    Public c_LED_blackout, c_LED_vision, c_LED_stage As Integer
    Public c_LED As Integer

    Public c_smoke_belimlight, c_smoke_PRLighting As Integer
    Public c_smoke_blackout, c_smoke_vision, c_smoke_stage As Integer
    Public c_smoke As Integer

    Public c_consoles_belimlight, c_consoles_PRLighting As Integer
    Public c_consoles_blackout, c_consoles_vision, c_consoles_stage As Integer
    Public c_consoles As Integer

    Public c_intercom_belimlight, c_intercom_PRLighting As Integer
    Public c_intercom_blackout, c_intercom_vision, c_intercom_stage As Integer
    Public c_intercom As Integer

    Public c_Light_tbl(7, 4) As Integer
    Public c_Light_sumTbl(7) As Integer

    '=================================      Address        ==================================================

    Public adr_movHeads_belimlight, adr_movHeads_PRLighting As String
    Public adr_movHeads_blackout, adr_movHeads_vision, adr_movHeads_stage As String
    Public adr_movHeads As String

    Public adr_strobes_belimlight, adr_strobes_PRLighting As String
    Public adr_strobes_blackout, adr_strobes_vision, adr_strobes_stage As String
    Public adr_strobes As String

    Public adr_blinders_belimlight, adr_blinders_PRLighting As String
    Public adr_blinders_blackout, adr_blinders_vision, adr_blinders_stage As String
    Public adr_blinders As String

    Public adr_arch_belimlight, adr_arch_PRLighting As String
    Public adr_arch_blackout, adr_arch_vision, adr_arch_stage As String
    Public adr_arch As String

    Public adr_LED_belimlight, adr_LED_PRLighting As String
    Public adr_LED_blackout, adr_LED_vision, adr_LED_stage As String
    Public adr_LED As String

    Public adr_smoke_belimlight, adr_smoke_PRLighting As String
    Public adr_smoke_blackout, adr_smoke_vision, adr_smoke_stage As String
    Public adr_smoke As String

    Public adr_consoles_belimlight, adr_consoles_PRLighting As String
    Public adr_consoles_blackout, adr_consoles_vision, adr_consoles_stage As String
    Public adr_consoles As String

    Public adr_intercom_belimlight, adr_intercom_PRLighting As String
    Public adr_intercom_blackout, adr_intercom_vision, adr_intercom_stage As String
    Public adr_intercom As String

    Public adr_Light_tbl(7, 4) As String
    Public adr_Light_sumTbl(7) As String

    '=================================      ExcelRange       =================================================

    Public rng_movHeads_belimlight, rng_movHeads_PRLighting As ExcelRange
    Public rng_movHeads_blackout, rng_movHeads_vision, rng_movHeads_stage As ExcelRange
    Public rng_movHeads As ExcelRange

    Public rng_strobes_belimlight, rng_strobes_PRLighting As ExcelRange
    Public rng_strobes_blackout, rng_strobes_vision, rng_strobes_stage As ExcelRange
    Public rng_strobes As ExcelRange

    Public rng_blinders_belimlight, rng_blinders_PRLighting As ExcelRange
    Public rng_blinders_blackout, rng_blinders_vision, rng_blinders_stage As ExcelRange
    Public rng_blinders As ExcelRange

    Public rng_arch_belimlight, rng_arch_PRLighting As ExcelRange
    Public rng_arch_blackout, rng_arch_vision, rng_arch_stage As ExcelRange
    Public rng_arch As ExcelRange

    Public rng_LED_belimlight, rng_LED_PRLighting As ExcelRange
    Public rng_LED_blackout, rng_LED_vision, rng_LED_stage As ExcelRange
    Public rng_LED As ExcelRange

    Public rng_smoke_belimlight, rng_smoke_PRLighting As ExcelRange
    Public rng_smoke_blackout, rng_smoke_vision, rng_smoke_stage As ExcelRange
    Public rng_smoke As ExcelRange

    Public rng_consoles_belimlight, rng_consoles_PRLighting As ExcelRange
    Public rng_consoles_blackout, rng_consoles_vision, rng_consoles_stage As ExcelRange
    Public rng_consoles As ExcelRange

    Public rng_intercom_belimlight, rng_intercom_PRLighting As ExcelRange
    Public rng_intercom_blackout, rng_intercom_vision, rng_intercom_stage As ExcelRange
    Public rng_intercom As ExcelRange

    Public rng_Light_tbl(7, 4) As ExcelRange
    Public rng_Light_sumTbl(7) As ExcelRange

    '=================================      Others       =================================================

    Public obj_excel, obj_excelFile As Object         '   Global vars to use in function "Save"

    Public sCompany() As String = {"belimlight", "PRLighting", "blackout", "vision", "stage"}
    Public selCompIndex As Integer = 0

    Public lblSumQty() As Object

    Public editMode() As String = {"Update", "Delete", "Add"}
    Public selEditModeIndex As Integer = 0
    Public selectedCategoryIndex As Integer



    '===================================================================================      
    '                === Load button ===
    '===================================================================================
    Private Sub btn_loadDB_Click(sender As Object, e As EventArgs) Handles btn_loadDB.Click



        loadDataBase()                      '   myFunctions.vb

        initLightWorksheets()               '   declarations.vb

        initLightTables()                   '   declarations.vb

        initDatatables()                    '   declarations.vb

        initLabels()                        '   declarations.vb


        '----------------------         Create datatables           ------------------------------
        '-----------------------------------------------------------------------------------------
        For i As Integer = 0 To cmb_category.Items.Count - 1

            For j As Integer = 0 To sCompany.Count - 1
                create_datatable(i, j)

            Next j
            'create_sumDatatable(i)
            create_sumDatatable_v2(i)
        Next i

        tabControl.SelectedIndex = 1
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

        Dim i As Integer

        i = cmb_category.SelectedIndex

        DGV_light.DataSource = dt_Lighting(i, 0)
        DGV_format(tbl_Lighting_tables(i, 0).Name, c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        DGV_light.Rows(0).Cells(0).Selected = True

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

        Dim i As Integer

        i = cmb_category.SelectedIndex

        DGV_light.DataSource = dt_Lighting(i, 1)
        DGV_format(tbl_Lighting_tables(i, 1).Name, c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        DGV_light.Rows(0).Cells(0).Selected = True
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

        Dim i As Integer

        i = cmb_category.SelectedIndex

        DGV_light.DataSource = dt_Lighting(i, 2)
        DGV_format(tbl_Lighting_tables(i, 2).Name, c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        DGV_light.Rows(0).Cells(0).Selected = True
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

        Dim i As Integer

        i = cmb_category.SelectedIndex

        DGV_light.DataSource = dt_Lighting(i, 3)
        DGV_format(tbl_Lighting_tables(i, 3).Name, c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        DGV_light.Rows(0).Cells(0).Selected = True
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

        Dim i As Integer

        i = cmb_category.SelectedIndex

        DGV_light.DataSource = dt_Lighting(i, 4)
        DGV_format(tbl_Lighting_tables(i, 4).Name, c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        DGV_light.Rows(0).Cells(0).Selected = True
        clearControls()

    End Sub

    '===================================================================================
    '             === Select page ===
    '===================================================================================
    Private Sub tabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabControl.SelectedIndexChanged
        cmb_category.SelectedIndex = 0
    End Sub
    '===================================================================================
    '             === Select category ===
    '===================================================================================
    Private Sub cmb_category_SelectedIndexChanged(sender As Object, e As EventArgs)
        clearControls()
        DGV_light.DataSource = Nothing
        btn_prev.Enabled = False
        btn_next.Enabled = False
    End Sub

    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Private Sub DGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_light.CellClick
        dgv_clickCell(sender, e)
        calcQuantity()
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

        initLightWorksheets()
        initLightTables()
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

        'Dim excelFile = New FileInfo(sFileName_DB)

        'ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        'Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

        'obj_excel = Excel                            '   Global vars to use in function "Save"
        'obj_excelFile = excelFile

        'initLightWorksheets()
        'initLightTables()
        'clearControls()
        unblockCompanyButtons()
        unblockEditButtons()

        btn_save.FlatStyle = FlatStyle.Standard

        'tabControl.SelectedIndex = 1

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

End Class
