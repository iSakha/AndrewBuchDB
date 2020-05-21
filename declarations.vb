Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO
Module declarations

    Sub initWorksheets(_tabIndex As Integer)

        Select Case _tabIndex
            Case 1

                mainForm.wsLight = {mainForm.obj_excel.Workbook.Worksheets(0),
                            mainForm.obj_excel.Workbook.Worksheets(1),
                            mainForm.obj_excel.Workbook.Worksheets(2),
                            mainForm.obj_excel.Workbook.Worksheets(3),
                            mainForm.obj_excel.Workbook.Worksheets(4),
                            mainForm.obj_excel.Workbook.Worksheets(5),
                            mainForm.obj_excel.Workbook.Worksheets(6),
                            mainForm.obj_excel.Workbook.Worksheets(7)}

            Case 2

                mainForm.wsScreen = {mainForm.obj_excel.Workbook.Worksheets(0),
                                mainForm.obj_excel.Workbook.Worksheets(1),
                                mainForm.obj_excel.Workbook.Worksheets(2),
                                mainForm.obj_excel.Workbook.Worksheets(3),
                                mainForm.obj_excel.Workbook.Worksheets(4),
                                mainForm.obj_excel.Workbook.Worksheets(5),
                                mainForm.obj_excel.Workbook.Worksheets(6),
                                mainForm.obj_excel.Workbook.Worksheets(7)}
        End Select

    End Sub

    Sub initExcelTables(_tabIndex As Integer)

        Select Case _tabIndex

            Case 1

                mainForm.tbl_Lighting_tables = {
            {mainForm.wsLight(0).Tables.Item("movHeads_belimlight"), mainForm.wsLight(0).Tables.
            Item("movHeads_PRLighting"), mainForm.wsLight(0).Tables.Item("movHeads_blackout"),
            mainForm.wsLight(0).Tables.Item("movHeads_vision"), mainForm.wsLight(0).Tables.Item("movHeads_stage")},
            {mainForm.wsLight(1).Tables.Item("strobes_belimlight"), mainForm.wsLight(1).Tables.
            Item("strobes_PRLighting"), mainForm.wsLight(1).Tables.Item("strobes_blackout"),
            mainForm.wsLight(1).Tables.Item("strobes_vision"), mainForm.wsLight(1).Tables.Item("strobes_stage")},
            {mainForm.wsLight(2).Tables.Item("blinders_belimlight"), mainForm.wsLight(2).Tables.
            Item("blinders_PRLighting"), mainForm.wsLight(2).Tables.Item("blinders_blackout"),
            mainForm.wsLight(2).Tables.Item("blinders_vision"), mainForm.wsLight(2).Tables.Item("blinders_stage")},
            {mainForm.wsLight(3).Tables.Item("arch_belimlight"), mainForm.wsLight(3).Tables.
            Item("arch_PRLighting"), mainForm.wsLight(3).Tables.Item("arch_blackout"),
            mainForm.wsLight(3).Tables.Item("arch_vision"), mainForm.wsLight(3).Tables.Item("arch_stage")},
            {mainForm.wsLight(4).Tables.Item("LED_belimlight"), mainForm.wsLight(4).Tables.
            Item("LED_PRLighting"), mainForm.wsLight(4).Tables.Item("LED_blackout"),
            mainForm.wsLight(4).Tables.Item("LED_vision"), mainForm.wsLight(4).Tables.Item("LED_stage")},
            {mainForm.wsLight(5).Tables.Item("smoke_belimlight"), mainForm.wsLight(5).Tables.
            Item("smoke_PRLighting"), mainForm.wsLight(5).Tables.Item("smoke_blackout"),
            mainForm.wsLight(5).Tables.Item("smoke_vision"), mainForm.wsLight(5).Tables.Item("smoke_stage")},
            {mainForm.wsLight(6).Tables.Item("consoles_belimlight"), mainForm.wsLight(6).
            Tables.Item("consoles_PRLighting"), mainForm.wsLight(6).Tables.Item("consoles_blackout"),
            mainForm.wsLight(6).Tables.Item("consoles_vision"), mainForm.wsLight(6).Tables.Item("consoles_stage")},
            {mainForm.wsLight(7).Tables.Item("intercom_belimlight"), mainForm.wsLight(7).Tables.
            Item("intercom_PRLighting"), mainForm.wsLight(7).Tables.Item("intercom_blackout"),
            mainForm.wsLight(7).Tables.Item("intercom_vision"), mainForm.wsLight(7).Tables.Item("intercom_stage")}
        }

                mainForm.tbl_Lighting_sumTables = {mainForm.wsLight(0).Tables.Item("movHeads_tbl"), mainForm.wsLight(1).Tables.
                    Item("strobes_tbl"), mainForm.wsLight(2).Tables.Item("blinders_tbl"), mainForm.wsLight(3).Tables.
                    Item("arch_tbl"), mainForm.wsLight(4).Tables.Item("LED_tbl"), mainForm.wsLight(5).Tables.
                    Item("smoke_tbl"), mainForm.wsLight(6).Tables.Item("consoles_tbl"), mainForm.wsLight(7).Tables.Item("intercom_tbl")}

            Case 2

                mainForm.tbl_Screen_tables = {
            {mainForm.wsScreen(0).Tables.Item("modules_belimlight"), mainForm.wsScreen(0).Tables.
            Item("modules_PRLighting"), mainForm.wsScreen(0).Tables.Item("modules_blackout"),
            mainForm.wsScreen(0).Tables.Item("modules_vision"), mainForm.wsScreen(0).Tables.Item("modules_stage")},
            {mainForm.wsScreen(1).Tables.Item("servers_belimlight"), mainForm.wsScreen(1).Tables.
            Item("servers_PRLighting"), mainForm.wsScreen(1).Tables.Item("servers_blackout"),
            mainForm.wsScreen(1).Tables.Item("servers_vision"), mainForm.wsScreen(1).Tables.Item("servers_stage")},
            {mainForm.wsScreen(2).Tables.Item("controllers1_belimlight"), mainForm.wsScreen(2).Tables.
            Item("controllers1_PRLighting"), mainForm.wsScreen(2).Tables.Item("controllers1_blackout"),
            mainForm.wsScreen(2).Tables.Item("controllers1_vision"), mainForm.wsScreen(2).Tables.Item("controllers1_stage")},
            {mainForm.wsScreen(3).Tables.Item("controllers2_belimlight"), mainForm.wsScreen(3).Tables.
            Item("controllers2_PRLighting"), mainForm.wsScreen(3).Tables.Item("controllers2_blackout"),
            mainForm.wsScreen(3).Tables.Item("controllers2_vision"), mainForm.wsScreen(3).Tables.Item("controllers2_stage")},
            {mainForm.wsScreen(4).Tables.Item("projectors_belimlight"), mainForm.wsScreen(4).Tables.
            Item("projectors_PRLighting"), mainForm.wsScreen(4).Tables.Item("projectors_blackout"),
            mainForm.wsScreen(4).Tables.Item("projectors_vision"), mainForm.wsScreen(4).Tables.Item("projectors_stage")},
            {mainForm.wsScreen(5).Tables.Item("construction_belimlight"), mainForm.wsScreen(5).Tables.
            Item("construction_PRLighting"), mainForm.wsScreen(5).Tables.Item("construction_blackout"),
            mainForm.wsScreen(5).Tables.Item("construction_vision"), mainForm.wsScreen(5).Tables.Item("construction_stage")},
            {mainForm.wsScreen(6).Tables.Item("lightingdesks_belimlight"), mainForm.wsScreen(6).
            Tables.Item("lightingdesks_PRLighting"), mainForm.wsScreen(6).Tables.Item("lightingdesks_blackout"),
            mainForm.wsScreen(6).Tables.Item("lightingdesks_vision"), mainForm.wsScreen(6).Tables.Item("lightingdesks_stage")},
            {mainForm.wsScreen(7).Tables.Item("cameras_belimlight"), mainForm.wsScreen(7).Tables.
            Item("cameras_PRLighting"), mainForm.wsScreen(7).Tables.Item("cameras_blackout"),
            mainForm.wsScreen(7).Tables.Item("cameras_vision"), mainForm.wsScreen(7).Tables.Item("cameras_stage")}}

                mainForm.tbl_Screen_sumTables = {mainForm.wsScreen(0).Tables.Item("modules_tbl"), mainForm.wsScreen(1).Tables.
                    Item("servers_tbl"), mainForm.wsScreen(2).Tables.Item("controllers1_tbl"), mainForm.wsScreen(3).Tables.
                    Item("controllers2_tbl"), mainForm.wsScreen(4).Tables.Item("projectors_tbl"), mainForm.wsScreen(5).Tables.
                    Item("construction_tbl"), mainForm.wsScreen(6).Tables.Item("lightingdesks_tbl"), mainForm.wsScreen(7).Tables.Item("cameras_tbl")}
        End Select
    End Sub

    Sub initLabels()
        mainForm.lblSumQty = {mainForm.lbl_qty_belimlight, mainForm.lbl_qty_PRLighting,
            mainForm.lbl_qty_blackout, mainForm.lbl_qty_vision, mainForm.lbl_qty_stage}
    End Sub


End Module
