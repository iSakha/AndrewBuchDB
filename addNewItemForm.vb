Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Public Class addNewItemForm
    Public txtName() As String
    Public txtQty() As String
    Private Sub addNewItemForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'txtName = {
        '    txt_belimlight1_addform.Text, txt_belimlight2_addform.Text, txt_belimlight3_addform.Text,
        '    txt_PRlighting1_addform.Text, txt_PRlighting2_addform.Text, txt_PRlighting3_addform.Text,
        '    txt_blackout1_addform.Text, txt_blackout2_addform.Text, txt_blackout3_addform.Text,
        '    txt_vision1_addform.Text, txt_vision2_addform.Text, txt_vision3_addform.Text,
        '    txt_stage1_addform.Text, txt_stage2_addform.Text, txt_stage3_addform.Text
        '    }

        'txtQty = {txt_qty_belimlight1_addform.Text, txt_qty_belimlight2_addform.Text, txt_qty_belimlight3_addform.Text,
        '    txt_qty_PRlighting1_addform.Text, txt_qty_PRlighting2_addform.Text, txt_qty_PRlighting3_addform.Text,
        '    txt_qty_blackout1_addform.Text, txt_qty_blackout2_addform.Text, txt_qty_blackout3_addform.Text,
        '    txt_qty_vision1_addform.Text, txt_qty_vision2_addform.Text, txt_qty_vision3_addform.Text,
        '    txt_qty_stage1_addform.Text, txt_qty_stage2_addform.Text, txt_qty_stage3_addform.Text
        '}
    End Sub

    Private Sub btn_save_addform_Click(sender As Object, e As EventArgs) Handles btn_save_addform.Click
        Dim sRow(4, 7) As String
        sRow = New String(4, 7) {}
        Dim i, j, rCount As Integer

        i = mainForm.selectedCategoryIndex

        sRow = {
            {txt_name_addform.Text, txt_qty_addform.Text,
            txt_belimlight1_addform.Text, txt_qty_belimlight1_addform.Text,
            txt_belimlight2_addform.Text, txt_qty_belimlight2_addform.Text,
            txt_belimlight3_addform.Text, txt_qty_belimlight3_addform.Text},
            {txt_name_addform.Text, txt_qty_addform.Text,
            txt_PRlighting1_addform.Text, txt_qty_PRlighting1_addform.Text,
            txt_PRlighting2_addform.Text, txt_qty_PRlighting2_addform.Text,
            txt_PRlighting3_addform.Text, txt_qty_PRlighting3_addform.Text},
            {txt_name_addform.Text, txt_qty_addform.Text,
            txt_blackout1_addform.Text, txt_qty_blackout1_addform.Text,
            txt_blackout2_addform.Text, txt_qty_blackout2_addform.Text,
            txt_blackout3_addform.Text, txt_qty_blackout3_addform.Text},
            {txt_name_addform.Text, txt_qty_addform.Text,
            txt_vision1_addform.Text, txt_qty_vision1_addform.Text,
            txt_vision2_addform.Text, txt_qty_vision2_addform.Text,
            txt_vision3_addform.Text, txt_qty_vision3_addform.Text},
            {txt_name_addform.Text, txt_qty_addform.Text,
            txt_stage1_addform.Text, txt_qty_stage1_addform.Text,
            txt_stage2_addform.Text, txt_qty_stage2_addform.Text,
            txt_stage3_addform.Text, txt_qty_stage3_addform.Text}
        }

        For j = 0 To mainForm.sCompany.Count - 1
            mainForm.dt_Lighting(i, j) = New DataTable
            create_datatable(mainForm.r_Light_tbl(i, j), mainForm.c_Light_tbl(i, j), mainForm.rng_Light_tbl(i, j),
                             mainForm.dt_Lighting(i, j), mainForm.tbl_Lighting_tables(i, j).Name)

            Dim row As DataRow

            row = mainForm.dt_Lighting(i, j).Rows.Add()
            rCount = mainForm.dt_Lighting(i, j).Rows.Count

            For k As Integer = 0 To 7
                If isEven(k) Then
                    row.Item(k + 1) = sRow(j, k)
                Else
                    row.Item(k + 1) = CInt(sRow(j, k))
                End If
            Next k

            row.Item(0) = CInt(mainForm.dt_Lighting(i, j).Rows(rCount - 2).Item(0)) + 1

        Next j

        'mainForm.DGV_light.DataSource = mainForm.dt_Lighting(i, 0)
        mainForm.selEditModeIndex = 2
        saveData(i, 777)


        clearControls()

        mainForm.btn_save.FlatStyle = FlatStyle.Standard
        unblockCompanyButtons()
        unblockEditButtons()

        Dim excelFile = New FileInfo(mainForm.sFileName_DB)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

        mainForm.obj_excel = Excel                            '   Global vars to use in function "Save"
        mainForm.obj_excelFile = excelFile

        initLightWorksheets()
        initLightTables()
        formatExcelTable(i, j)

    End Sub
    Private Sub cmb_category_addform_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_category_addform.SelectedIndexChanged
        mainForm.selectedCategoryIndex = cmb_category_addform.SelectedIndex
    End Sub
    Private Sub btn_close_addform_Click(sender As Object, e As EventArgs) Handles btn_close_addform.Click
        Me.Close()
    End Sub


End Class