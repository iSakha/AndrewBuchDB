Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Public Class mainForm

    Public sDir_DB As String
    Public sFileName_DB As String

    Public wsMovHeads As ExcelWorksheet

    Public tbl_Light_Collection As ExcelTableCollection

    Public tbl_movHeads_belimlight, tbl_movHeads_PRLighting As ExcelTable
    Public tbl_movHeads_blackout, tbl_movHeads_vision As ExcelTable

    Public dt_movHeads_belimlight, dt_movHeads_PRLighting As New DataTable()
    Public dt_movHeads_blackout, dt_movHeads_vision As New DataTable()

    Public r_movHeads_belimlight, r_movHeads_PRLighting As Integer
    Public r_movHeads_blackout, r_movHeads_vision As Integer

    Public c_movHeads_belimlight, c_movHeads_PRLighting As Integer
    Public c_movHeads_blackout, c_movHeads_vision As Integer

    Public adr_movHeads_belimlight, adr_movHeads_PRlighting As String
    Public adr_movHeads_blackout, adr_movHeads_vision As String

    Public rng_movHeads_belimlight, rng_movHeads_PRlighting As ExcelRange
    Public rng_movHeads_blackout, rng_movHeads_vision As ExcelRange

    Public obj_excelLight, obj_excelFileLight As Object         '   Global vars to use in function "Save"

    '===================================================================================      
    '                === Load button ===
    '===================================================================================
    Private Sub btn_loadDB_Click(sender As Object, e As EventArgs) Handles btn_loadDB.Click

        sDir_DB = Directory.GetCurrentDirectory()

        OFD.InitialDirectory = sDir_DB
        OFD.Title = "Select .omdb file"

        If OFD.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            sFileName_DB = OFD.FileName

            Dim excelFile = New FileInfo(sFileName_DB)

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Dim Excel As ExcelPackage = New ExcelPackage(excelFile)

            obj_excelLight = Excel                            '   Global vars to use in function "Save"
            obj_excelFileLight = excelFile


            wsMovHeads = Excel.Workbook.Worksheets(0)

            tbl_Light_Collection = wsMovHeads.Tables

            tbl_movHeads_belimlight = tbl_Light_Collection.Item("movHeads_belimlight")
            tbl_movHeads_PRLighting = tbl_Light_Collection.Item("movHeads_PRlighting")
            tbl_movHeads_blackout = tbl_Light_Collection.Item("movHeads_blackout")
            tbl_movHeads_vision = tbl_Light_Collection.Item("movHeads_vision")

            r_movHeads_belimlight = tbl_movHeads_belimlight.Address.Rows
            r_movHeads_PRLighting = tbl_movHeads_PRLighting.Address.Rows
            r_movHeads_blackout = tbl_movHeads_blackout.Address.Rows
            r_movHeads_vision = tbl_movHeads_vision.Address.Rows

            c_movHeads_belimlight = tbl_movHeads_belimlight.Address.Columns
            c_movHeads_PRLighting = tbl_movHeads_PRLighting.Address.Columns
            c_movHeads_blackout = tbl_movHeads_blackout.Address.Columns
            c_movHeads_vision = tbl_movHeads_vision.Address.Columns

            adr_movHeads_belimlight = tbl_movHeads_belimlight.Address.Address
            adr_movHeads_PRlighting = tbl_movHeads_PRLighting.Address.Address
            adr_movHeads_blackout = tbl_movHeads_blackout.Address.Address
            adr_movHeads_vision = tbl_movHeads_vision.Address.Address

            rng_movHeads_belimlight = wsMovHeads.Cells(adr_movHeads_belimlight)
            rng_movHeads_PRlighting = wsMovHeads.Cells(adr_movHeads_PRlighting)
            rng_movHeads_blackout = wsMovHeads.Cells(adr_movHeads_blackout)
            rng_movHeads_vision = wsMovHeads.Cells(adr_movHeads_vision)

            tabControl.SelectedIndex = 1

        End If
    End Sub

    '===================================================================================      
    '                === Belimlight button ===
    '===================================================================================
    Private Sub btn_belIm_Click(sender As Object, e As EventArgs) Handles btn_belIm.Click

        create_datatable(r_movHeads_belimlight, c_movHeads_belimlight, rng_movHeads_belimlight, dt_movHeads_belimlight, "belimlight")
        DGV.DataSource = dt_movHeads_belimlight
        DGV_format("belimlight")

    End Sub

    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Private Sub DGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellClick
        dgv_clickCell(sender, e)
    End Sub

End Class
