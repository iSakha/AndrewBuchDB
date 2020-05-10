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
    Public tbl_movHeads_blackout, tbl_movHeads_vision As ExcelTable

    Public tbl_strobes_belimlight, tbl_strobes_PRLighting As ExcelTable
    Public tbl_strobes_blackout, tbl_strobes_vision As ExcelTable

    Public tbl_blinders_belimlight, tbl_blinders_PRLighting As ExcelTable
    Public tbl_blinders_blackout, tbl_blinders_vision As ExcelTable

    Public tbl_arch_belimlight, tbl_arch_PRLighting As ExcelTable
    Public tbl_arch_blackout, tbl_arch_vision As ExcelTable

    Public tbl_LED_belimlight, tbl_LED_PRLighting As ExcelTable
    Public tbl_LED_blackout, tbl_LED_vision As ExcelTable

    Public tbl_smoke_belimlight, tbl_smoke_PRLighting As ExcelTable
    Public tbl_smoke_blackout, tbl_smoke_vision As ExcelTable

    Public tbl_consoles_belimlight, tbl_consoles_PRLighting As ExcelTable
    Public tbl_consoles_blackout, tbl_consoles_vision As ExcelTable

    Public tbl_intercom_belimlight, tbl_intercom_PRLighting As ExcelTable
    Public tbl_intercom_blackout, tbl_intercom_vision As ExcelTable

    Public tbl_Lighting_tables(7, 3) As ExcelTable

    '=================================  DataTable  ========================================================

    Public dt_movHeads_belimlight, dt_movHeads_PRLighting As New DataTable()
    Public dt_movHeads_blackout, dt_movHeads_vision As New DataTable()

    Public dt_strobes_belimlight, dt_strobes_PRLighting As New DataTable()
    Public dt_strobes_blackout, dt_strobes_vision As New DataTable()

    Public dt_blinders_belimlight, dt_blinders_PRLighting As New DataTable()
    Public dt_blinders_blackout, dt_blinders_vision As New DataTable()

    Public dt_arch_belimlight, dt_arch_PRLighting As New DataTable()
    Public dt_arch_blackout, dt_arch_vision As New DataTable()

    Public dt_LED_belimlight, dt_LED_PRLighting As New DataTable()
    Public dt_LED_blackout, dt_LED_vision As New DataTable()

    Public dt_smoke_belimlight, dt_smoke_PRLighting As New DataTable()
    Public dt_smoke_blackout, dt_smoke_vision As New DataTable()

    Public dt_consoles_belimlight, dt_consoles_PRLighting As New DataTable()
    Public dt_consoles_blackout, dt_consoles_vision As New DataTable()

    Public dt_intercom_belimlight, dt_intercom_PRLighting As New DataTable()
    Public dt_intercom_blackout, dt_intercom_vision As New DataTable()

    Public dt_Lighting(7, 3) As Object

    '=================================  Rows and columns  ==================================================

    Public r_movHeads_belimlight, r_movHeads_PRLighting As Integer
    Public r_movHeads_blackout, r_movHeads_vision As Integer

    Public r_strobes_belimlight, r_strobes_PRLighting As Integer
    Public r_strobes_blackout, r_strobes_vision As Integer

    Public r_blinders_belimlight, r_blinders_PRLighting As Integer
    Public r_blinders_blackout, r_blinders_vision As Integer

    Public r_arch_belimlight, r_arch_PRLighting As Integer
    Public r_arch_blackout, r_arch_vision As Integer

    Public r_LED_belimlight, r_LED_PRLighting As Integer
    Public r_LED_blackout, r_LED_vision As Integer

    Public r_smoke_belimlight, r_smoke_PRLighting As Integer
    Public r_smoke_blackout, r_smoke_vision As Integer

    Public r_consoles_belimlight, r_consoles_PRLighting As Integer
    Public r_consoles_blackout, r_consoles_vision As Integer

    Public r_intercom_belimlight, r_intercom_PRLighting As Integer
    Public r_intercom_blackout, r_intercom_vision As Integer

    Public r_Light_tbl(7, 3) As Integer

    Public c_movHeads_belimlight, c_movHeads_PRLighting As Integer
    Public c_movHeads_blackout, c_movHeads_vision As Integer

    Public c_strobes_belimlight, c_strobes_PRLighting As Integer
    Public c_strobes_blackout, c_strobes_vision As Integer

    Public c_blinders_belimlight, c_blinders_PRLighting As Integer
    Public c_blinders_blackout, c_blinders_vision As Integer

    Public c_arch_belimlight, c_arch_PRLighting As Integer
    Public c_arch_blackout, c_arch_vision As Integer

    Public c_LED_belimlight, c_LED_PRLighting As Integer
    Public c_LED_blackout, c_LED_vision As Integer

    Public c_smoke_belimlight, c_smoke_PRLighting As Integer
    Public c_smoke_blackout, c_smoke_vision As Integer

    Public c_consoles_belimlight, c_consoles_PRLighting As Integer
    Public c_consoles_blackout, c_consoles_vision As Integer

    Public c_intercom_belimlight, c_intercom_PRLighting As Integer
    Public c_intercom_blackout, c_intercom_vision As Integer

    Public c_Light_tbl(7, 3) As Integer

    '=================================      Address        ==================================================

    Public adr_movHeads_belimlight, adr_movHeads_PRLighting As String
    Public adr_movHeads_blackout, adr_movHeads_vision As String

    Public adr_strobes_belimlight, adr_strobes_PRLighting As String
    Public adr_strobes_blackout, adr_strobes_vision As String

    Public adr_blinders_belimlight, adr_blinders_PRLighting As String
    Public adr_blinders_blackout, adr_blinders_vision As String

    Public adr_arch_belimlight, adr_arch_PRLighting As String
    Public adr_arch_blackout, adr_arch_vision As String

    Public adr_LED_belimlight, adr_LED_PRLighting As String
    Public adr_LED_blackout, adr_LED_vision As String

    Public adr_smoke_belimlight, adr_smoke_PRLighting As String
    Public adr_smoke_blackout, adr_smoke_vision As String

    Public adr_consoles_belimlight, adr_consoles_PRLighting As String
    Public adr_consoles_blackout, adr_consoles_vision As String

    Public adr_intercom_belimlight, adr_intercom_PRLighting As String
    Public adr_intercom_blackout, adr_intercom_vision As String

    Public adr_Light_tbl(7, 3) As String

    '=================================      ExcelRange       =================================================

    Public rng_movHeads_belimlight, rng_movHeads_PRLighting As ExcelRange
    Public rng_movHeads_blackout, rng_movHeads_vision As ExcelRange

    Public rng_strobes_belimlight, rng_strobes_PRLighting As ExcelRange
    Public rng_strobes_blackout, rng_strobes_vision As ExcelRange

    Public rng_blinders_belimlight, rng_blinders_PRLighting As ExcelRange
    Public rng_blinders_blackout, rng_blinders_vision As ExcelRange

    Public rng_arch_belimlight, rng_arch_PRLighting As ExcelRange
    Public rng_arch_blackout, rng_arch_vision As ExcelRange

    Public rng_LED_belimlight, rng_LED_PRLighting As ExcelRange
    Public rng_LED_blackout, rng_LED_vision As ExcelRange

    Public rng_smoke_belimlight, rng_smoke_PRLighting As ExcelRange
    Public rng_smoke_blackout, rng_smoke_vision As ExcelRange

    Public rng_consoles_belimlight, rng_consoles_PRLighting As ExcelRange
    Public rng_consoles_blackout, rng_consoles_vision As ExcelRange

    Public rng_intercom_belimlight, rng_intercom_PRLighting As ExcelRange
    Public rng_intercom_blackout, rng_intercom_vision As ExcelRange

    Public rng_Light_tbl(7, 3) As ExcelRange

    '=================================      Others       =================================================

    Public obj_excel, obj_excelFile As Object         '   Global vars to use in function "Save"

    Public selectedCompany() As String = {"belimlight", "PRLighting", "blackout", "vision"}
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

            obj_excel = Excel                            '   Global vars to use in function "Save"
            obj_excelFile = excelFile

            initLightWorksheets()
            initLightTables()



            'tbl_Light_Collection = wsLight(0).Tables
            'initMovHeads()
            'tbl_Light_Collection = wsLight(1).Tables
            'initStrobes()
            'tbl_Light_Collection = wsLight(2).Tables
            'initBlinders()
            'tbl_Light_Collection = wsLight(3).Tables
            'initArch()
            'tbl_Light_Collection = wsLight(4).Tables
            'initLED()
            'tbl_Light_Collection = wsLight(5).Tables
            'initSmoke()
            'tbl_Light_Collection = wsLight(6).Tables
            'initConsoles()
            'tbl_Light_Collection = wsLight(7).Tables
            'initIntercom()

            tabControl.SelectedIndex = 1

        End If
    End Sub

    '===================================================================================      
    '                === Belimlight button ===
    '===================================================================================
    Private Sub btn_belIm_Click(sender As Object, e As EventArgs) Handles btn_belIm.Click

        selComp = selectedCompany(0)

        Dim c As Color = Color.FromArgb(252, 228, 214)

        Dim i As Integer

        i = cmb_category.SelectedIndex
        create_datatable(r_Light_tbl(i, 0), c_Light_tbl(i, 0), rng_Light_tbl(i, 0), dt_Lighting(i, 0), tbl_Lighting_tables(i, 0).Name)
        DGV.DataSource = dt_Lighting(i, 0)
        DGV_format(tbl_Lighting_tables(i, 0).Name, c)

        rtb_fixtureName.BackColor = c
        rtb_FirstName.BackColor = c
        rtb_SecondName.BackColor = c
        rtb_ThirdName.BackColor = c

        'Select Case cmb_category.SelectedIndex

        '    Case 0

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_movHeads_belimlight = New DataTable
        '        create_datatable(r_movHeads_belimlight, c_movHeads_belimlight, rng_movHeads_belimlight, dt_movHeads_belimlight, "movHeads_belimlight")
        '        DGV.DataSource = dt_movHeads_belimlight
        '        DGV_format("movHeads_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        '    Case 1

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_strobes_belimlight = New DataTable
        '        create_datatable(r_strobes_belimlight, c_strobes_belimlight, rng_strobes_belimlight, dt_strobes_belimlight, "strobes_belimlight")
        '        DGV.DataSource = dt_strobes_belimlight
        '        DGV_format("strobes_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        '    Case 2

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_blinders_belimlight = New DataTable
        '        create_datatable(r_blinders_belimlight, c_blinders_belimlight, rng_blinders_belimlight, dt_blinders_belimlight, "blinders_belimlight")
        '        DGV.DataSource = dt_blinders_belimlight
        '        DGV_format("blinders_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        '    Case 3

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_arch_belimlight = New DataTable
        '        create_datatable(r_arch_belimlight, c_arch_belimlight, rng_arch_belimlight, dt_arch_belimlight, "arch_belimlight")
        '        DGV.DataSource = dt_arch_belimlight
        '        DGV_format("arch_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        '    Case 4

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_LED_belimlight = New DataTable
        '        create_datatable(r_LED_belimlight, c_LED_belimlight, rng_LED_belimlight, dt_LED_belimlight, "LED_belimlight")
        '        DGV.DataSource = dt_LED_belimlight
        '        DGV_format("LED_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        '    Case 5

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_smoke_belimlight = New DataTable
        '        create_datatable(r_smoke_belimlight, c_smoke_belimlight, rng_smoke_belimlight, dt_smoke_belimlight, "smoke_belimlight")
        '        DGV.DataSource = dt_smoke_belimlight
        '        DGV_format("smoke_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        '    Case 6

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_consoles_belimlight = New DataTable
        '        create_datatable(r_consoles_belimlight, c_consoles_belimlight, rng_consoles_belimlight, dt_consoles_belimlight, "consoles_belimlight")
        '        DGV.DataSource = dt_consoles_belimlight
        '        DGV_format("consoles_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        '    Case 7

        '        Dim c As Color = Color.FromArgb(252, 228, 214)
        '        dt_intercom_belimlight = New DataTable
        '        create_datatable(r_intercom_belimlight, c_intercom_belimlight, rng_intercom_belimlight, dt_intercom_belimlight, "intercom_belimlight")
        '        DGV.DataSource = dt_intercom_belimlight
        '        DGV_format("intercom_belimlight", c)

        '        rtb_fixtureName.BackColor = c
        '        rtb_FirstName.BackColor = c
        '        rtb_SecondName.BackColor = c
        '        rtb_ThirdName.BackColor = c

        'End Select

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
                create_datatable(r_movHeads_PRLighting, c_movHeads_PRLighting, rng_movHeads_PRLighting, dt_movHeads_PRLighting, "movHeads_PRLighting")
                DGV.DataSource = dt_movHeads_PRLighting
                DGV_format("PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 1

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_strobes_PRLighting = New DataTable
                create_datatable(r_strobes_PRLighting, c_strobes_PRLighting, rng_strobes_PRLighting, dt_strobes_PRLighting, "strobes_PRLighting")
                DGV.DataSource = dt_strobes_PRLighting
                DGV_format("strobes_PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 2

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_blinders_PRLighting = New DataTable
                create_datatable(r_blinders_PRLighting, c_blinders_PRLighting, rng_blinders_PRLighting, dt_blinders_PRLighting, "blinders_PRLighting")
                DGV.DataSource = dt_blinders_PRLighting
                DGV_format("blinders_PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 3

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_arch_PRLighting = New DataTable
                create_datatable(r_arch_PRLighting, c_arch_PRLighting, rng_arch_PRLighting, dt_arch_PRLighting, "arch_PRLighting")
                DGV.DataSource = dt_arch_PRLighting
                DGV_format("arch_PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 4

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_LED_PRLighting = New DataTable
                create_datatable(r_LED_PRLighting, c_LED_PRLighting, rng_LED_PRLighting, dt_LED_PRLighting, "LED_PRLighting")
                DGV.DataSource = dt_LED_PRLighting
                DGV_format("LED_PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 5

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_smoke_PRLighting = New DataTable
                create_datatable(r_smoke_PRLighting, c_smoke_PRLighting, rng_smoke_PRLighting, dt_smoke_PRLighting, "smoke_PRLighting")
                DGV.DataSource = dt_smoke_PRLighting
                DGV_format("smoke_PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 6

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_consoles_PRLighting = New DataTable
                create_datatable(r_consoles_PRLighting, c_consoles_PRLighting, rng_consoles_PRLighting, dt_consoles_PRLighting, "consoles_PRLighting")
                DGV.DataSource = dt_consoles_PRLighting
                DGV_format("consoles_PRLighting", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 7

                Dim c As Color = Color.FromArgb(221, 235, 247)
                dt_intercom_PRLighting = New DataTable
                create_datatable(r_intercom_PRLighting, c_intercom_PRLighting, rng_intercom_PRLighting, dt_intercom_PRLighting, "intercom_PRLighting")
                DGV.DataSource = dt_intercom_PRLighting
                DGV_format("intercom_PRLighting", c)

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

            Case 2

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_blinders_blackout = New DataTable
                create_datatable(r_blinders_blackout, c_blinders_blackout, rng_blinders_blackout, dt_blinders_blackout, "blinders_movHeads_blackout")
                DGV.DataSource = dt_blinders_blackout
                DGV_format("blinders_blackout", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 3

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_arch_blackout = New DataTable
                create_datatable(r_arch_blackout, c_arch_blackout, rng_arch_blackout, dt_arch_blackout, "arch_movHeads_blackout")
                DGV.DataSource = dt_arch_blackout
                DGV_format("arch_blackout", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 4

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_LED_blackout = New DataTable
                create_datatable(r_LED_blackout, c_LED_blackout, rng_LED_blackout, dt_LED_blackout, "LED_movHeads_blackout")
                DGV.DataSource = dt_LED_blackout
                DGV_format("LED_blackout", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 5

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_smoke_blackout = New DataTable
                create_datatable(r_smoke_blackout, c_smoke_blackout, rng_smoke_blackout, dt_smoke_blackout, "smoke_movHeads_blackout")
                DGV.DataSource = dt_smoke_blackout
                DGV_format("smoke_blackout", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 6

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_consoles_blackout = New DataTable
                create_datatable(r_consoles_blackout, c_consoles_blackout, rng_consoles_blackout, dt_consoles_blackout, "consoles_movHeads_blackout")
                DGV.DataSource = dt_consoles_blackout
                DGV_format("consoles_blackout", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 7

                Dim c As Color = Color.FromArgb(237, 237, 237)
                dt_intercom_blackout = New DataTable
                create_datatable(r_intercom_blackout, c_intercom_blackout, rng_intercom_blackout, dt_intercom_blackout, "intercom_movHeads_blackout")
                DGV.DataSource = dt_intercom_blackout
                DGV_format("intercom_blackout", c)

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

            Case 2

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_blinders_vision = New DataTable
                create_datatable(r_blinders_vision, c_blinders_vision, rng_blinders_vision, dt_blinders_vision, "blinders_vision")
                DGV.DataSource = dt_blinders_vision
                DGV_format("blinders_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 3

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_arch_vision = New DataTable
                create_datatable(r_arch_vision, c_arch_vision, rng_arch_vision, dt_arch_vision, "arch_vision")
                DGV.DataSource = dt_arch_vision
                DGV_format("arch_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 4

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_LED_vision = New DataTable
                create_datatable(r_LED_vision, c_LED_vision, rng_LED_vision, dt_LED_vision, "LED_vision")
                DGV.DataSource = dt_LED_vision
                DGV_format("LED_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 5

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_smoke_vision = New DataTable
                create_datatable(r_smoke_vision, c_smoke_vision, rng_smoke_vision, dt_smoke_vision, "smoke_vision")
                DGV.DataSource = dt_smoke_vision
                DGV_format("smoke_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 6

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_consoles_vision = New DataTable
                create_datatable(r_consoles_vision, c_consoles_vision, rng_consoles_vision, dt_consoles_vision, "consoles_vision")
                DGV.DataSource = dt_consoles_vision
                DGV_format("consoles_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

            Case 6

                Dim c As Color = Color.FromArgb(226, 239, 218)
                dt_intercom_vision = New DataTable
                create_datatable(r_intercom_vision, c_intercom_vision, rng_intercom_vision, dt_intercom_vision, "intercom_vision")
                DGV.DataSource = dt_intercom_vision
                DGV_format("intercom_vision", c)

                rtb_fixtureName.BackColor = c
                rtb_FirstName.BackColor = c
                rtb_SecondName.BackColor = c
                rtb_ThirdName.BackColor = c

        End Select

    End Sub

    '===================================================================================
    '             === Select category ===
    '===================================================================================
    Private Sub cmb_category_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_category.SelectedIndexChanged

    End Sub


    '===================================================================================
    '             === CellClick on DGV ===
    '===================================================================================
    Private Sub DGV_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellClick
        dgv_clickCell(sender, e)
    End Sub

End Class
