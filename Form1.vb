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

    Public selectedCompany() As String = {"belimlight", "PRlighting", "blackout", "vision"}
    Public selComp As String = ""

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

        selComp = selectedCompany(0)

        Dim c As Color = Color.FromArgb(252, 228, 214)
        dt_movHeads_belimlight = New DataTable
        create_datatable(r_movHeads_belimlight, c_movHeads_belimlight, rng_movHeads_belimlight, dt_movHeads_belimlight, "belimlight")
        DGV.DataSource = dt_movHeads_belimlight
        DGV_format("belimlight", c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

    End Sub
    '===================================================================================      
    '                === PRLighting button ===
    '===================================================================================

    Private Sub btn_prLight_Click(sender As Object, e As EventArgs) Handles btn_prLight.Click

        selComp = selectedCompany(1)

        Dim c As Color = Color.FromArgb(221, 235, 247)
        dt_movHeads_PRLighting = New DataTable
        create_datatable(r_movHeads_PRLighting, c_movHeads_PRLighting, rng_movHeads_PRlighting, dt_movHeads_PRLighting, "PRLighting")
        DGV.DataSource = dt_movHeads_PRLighting
        DGV_format("PRLighting", c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

    End Sub
    '===================================================================================      
    '                === Blackout button ===
    '===================================================================================
    Private Sub btn_blackOut_Click(sender As Object, e As EventArgs) Handles btn_blackOut.Click

        selComp = selectedCompany(2)

        Dim c As Color = Color.FromArgb(237, 237, 237)
        dt_movHeads_blackout = New DataTable
        create_datatable(r_movHeads_blackout, c_movHeads_blackout, rng_movHeads_blackout, dt_movHeads_blackout, "blackout")
        DGV.DataSource = dt_movHeads_blackout
        DGV_format("blackout", c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

    End Sub


    '===================================================================================      
    '                === Vision button ===
    '===================================================================================
    Private Sub btn_vision_Click(sender As Object, e As EventArgs) Handles btn_vision.Click

        selComp = selectedCompany(0)

        Dim c As Color = Color.FromArgb(226, 239, 218)
        dt_movHeads_vision = New DataTable
        create_datatable(r_movHeads_vision, c_movHeads_vision, rng_movHeads_vision, dt_movHeads_vision, "vision")
        DGV.DataSource = dt_movHeads_vision
        DGV_format("vision", c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

    End Sub


    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Private Sub DGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellClick
        dgv_clickCell(sender, e)
    End Sub

End Class
