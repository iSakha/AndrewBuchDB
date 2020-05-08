Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Public Class mainForm

    Public sDir_DB As String
    Public sFileName_DB As String

    Public wsMovHeads As ExcelWorksheet
    Public wsStrobes As ExcelWorksheet

    Public tbl_Light_Collection As ExcelTableCollection

    Public tbl_movHeads_belimlight, tbl_movHeads_PRLighting As ExcelTable
    Public tbl_movHeads_blackout, tbl_movHeads_vision As ExcelTable

    Public tbl_strobes_belimlight, tbl_strobes_PRLighting As ExcelTable
    Public tbl_strobes_blackout, tbl_strobes_vision As ExcelTable

    Public dt_movHeads_belimlight, dt_movHeads_PRLighting As New DataTable()
    Public dt_movHeads_blackout, dt_movHeads_vision As New DataTable()

    Public dt_strobes_belimlight, dt_strobes_PRLighting As New DataTable()
    Public dt_strobes_blackout, dt_strobes_vision As New DataTable()

    Public r_movHeads_belimlight, r_movHeads_PRLighting As Integer
    Public r_movHeads_blackout, r_movHeads_vision As Integer

    Public r_strobes_belimlight, r_strobes_PRLighting As Integer
    Public r_strobes_blackout, r_strobes_vision As Integer

    Public c_movHeads_belimlight, c_movHeads_PRLighting As Integer
    Public c_movHeads_blackout, c_movHeads_vision As Integer

    Public c_strobes_belimlight, c_strobes_PRLighting As Integer
    Public c_strobes_blackout, c_strobes_vision As Integer

    Public adr_movHeads_belimlight, adr_movHeads_PRlighting As String
    Public adr_movHeads_blackout, adr_movHeads_vision As String

    Public adr_strobes_belimlight, adr_strobes_PRlighting As String
    Public adr_strobes_blackout, adr_strobes_vision As String

    Public rng_movHeads_belimlight, rng_movHeads_PRlighting As ExcelRange
    Public rng_movHeads_blackout, rng_movHeads_vision As ExcelRange

    Public rng_strobes_belimlight, rng_strobes_PRlighting As ExcelRange
    Public rng_strobes_blackout, rng_strobes_vision As ExcelRange

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
            wsStrobes = Excel.Workbook.Worksheets(1)

            tbl_Light_Collection = wsMovHeads.Tables
            initMovHeads()
            tbl_Light_Collection = wsStrobes.Tables
            initStrobes()
            tabControl.SelectedIndex = 1

        End If
    End Sub

    '===================================================================================      
    '                === Belimlight button ===
    '===================================================================================
    Private Sub btn_belIm_Click(sender As Object, e As EventArgs) Handles btn_belIm.Click

        selComp = selectedCompany(0)

        Select Case cmb_category.SelectedIndex

            Case 0

                Dim c As Color = Color.FromArgb(252, 228, 214)
                dt_movHeads_belimlight = New DataTable
                create_datatable(r_movHeads_belimlight, c_movHeads_belimlight, rng_movHeads_belimlight, dt_movHeads_belimlight, "movHeads_belimlight")
                DGV.DataSource = dt_movHeads_belimlight
                DGV_format("movHeads_belimlight", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 1

                Dim c As Color = Color.FromArgb(252, 228, 214)
                dt_strobes_belimlight = New DataTable
                create_datatable(r_strobes_belimlight, c_strobes_belimlight, rng_strobes_belimlight, dt_strobes_belimlight, "strobes_belimlight")
                DGV.DataSource = dt_strobes_belimlight
                DGV_format("strobes_belimlight", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

        End Select

    End Sub
    '===================================================================================      
    '                === PRLighting button ===
    '===================================================================================

    Private Sub btn_prLight_Click(sender As Object, e As EventArgs) Handles btn_prLight.Click

        selComp = selectedCompany(1)

        Select Case cmb_category.SelectedIndex

            Case 0

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_movHeads_PRLighting = New DataTable
                create_datatable(r_movHeads_PRLighting, c_movHeads_PRLighting, rng_movHeads_PRlighting, dt_movHeads_PRLighting, "movHeads_PRLighting")
                DGV.DataSource = dt_movHeads_PRLighting
                DGV_format("PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 1

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_strobes_PRLighting = New DataTable
                create_datatable(r_strobes_PRLighting, c_strobes_PRLighting, rng_strobes_PRlighting, dt_strobes_PRLighting, "strobes_PRLighting")
                DGV.DataSource = dt_strobes_PRLighting
                DGV_format("strobes_PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

        End Select

    End Sub
    '===================================================================================      
    '                === Blackout button ===
    '===================================================================================
    Private Sub btn_blackOut_Click(sender As Object, e As EventArgs) Handles btn_blackOut.Click

        selComp = selectedCompany(2)

        Select Case cmb_category.SelectedIndex

            Case 0

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_movHeads_blackout = New DataTable
                create_datatable(r_movHeads_blackout, c_movHeads_blackout, rng_movHeads_blackout, dt_movHeads_blackout, "movHeads_blackout")
                DGV.DataSource = dt_movHeads_blackout
                DGV_format("movHeads_blackout", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 1

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_strobes_blackout = New DataTable
                create_datatable(r_strobes_blackout, c_strobes_blackout, rng_strobes_blackout, dt_strobes_blackout, "strobes_movHeads_blackout")
                DGV.DataSource = dt_strobes_blackout
                DGV_format("strobes_blackout", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

        End Select

    End Sub


    '===================================================================================      
    '                === Vision button ===
    '===================================================================================
    Private Sub btn_vision_Click(sender As Object, e As EventArgs) Handles btn_vision.Click

        selComp = selectedCompany(0)

        Select Case cmb_category.SelectedIndex

            Case 0

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_movHeads_vision = New DataTable
                create_datatable(r_movHeads_vision, c_movHeads_vision, rng_movHeads_vision, dt_movHeads_vision, "movHeads_vision")
                DGV.DataSource = dt_movHeads_vision
                DGV_format("movHeads_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 1

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_strobes_vision = New DataTable
                create_datatable(r_strobes_vision, c_strobes_vision, rng_strobes_vision, dt_strobes_vision, "strobes_vision")
                DGV.DataSource = dt_strobes_vision
                DGV_format("strobes_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

        End Select

    End Sub


    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Private Sub DGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellClick
        dgv_clickCell(sender, e)
    End Sub

End Class
