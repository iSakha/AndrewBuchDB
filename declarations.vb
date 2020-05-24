Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO
Module declarations

    Sub initWorksheets(_tabIndex As Integer)

        Select Case _tabIndex
            Case 1

                'mainForm.wsLight = {mainForm.obj_excel.Workbook.Worksheets(0),
                '            mainForm.obj_excel.Workbook.Worksheets(1),
                '            mainForm.obj_excel.Workbook.Worksheets(2),
                '            mainForm.obj_excel.Workbook.Worksheets(3),
                '            mainForm.obj_excel.Workbook.Worksheets(4),
                '            mainForm.obj_excel.Workbook.Worksheets(5),
                '            mainForm.obj_excel.Workbook.Worksheets(6),
                '            mainForm.obj_excel.Workbook.Worksheets(7)}

            Case 2


        End Select

    End Sub

    Sub initExcelTables(_tabIndex As Integer)


        mainForm.tbl_Lighting_sumTables = {mainForm.wsLight(0).Tables.Item("movHeads_tbl"), mainForm.wsLight(1).Tables.
                    Item("strobes_tbl"), mainForm.wsLight(2).Tables.Item("blinders_tbl"), mainForm.wsLight(3).Tables.
                    Item("arch_tbl"), mainForm.wsLight(4).Tables.Item("LED_tbl"), mainForm.wsLight(5).Tables.
                    Item("smoke_tbl"), mainForm.wsLight(6).Tables.Item("consoles_tbl"), mainForm.wsLight(7).Tables.Item("intercom_tbl")}

        mainForm.tbl_Screen_sumTables = {mainForm.wsScreen(0).Tables.Item("modules_tbl"), mainForm.wsScreen(1).Tables.
                    Item("servers_tbl"), mainForm.wsScreen(2).Tables.Item("controllers1_tbl"), mainForm.wsScreen(3).Tables.
                    Item("controllers2_tbl"), mainForm.wsScreen(4).Tables.Item("projectors_tbl"), mainForm.wsScreen(5).Tables.
                    Item("construction_tbl"), mainForm.wsScreen(6).Tables.Item("lightingdesks_tbl"), mainForm.wsScreen(7).Tables.Item("cameras_tbl")}



    End Sub

    Sub initLabels()
        mainForm.lblSumQty = {mainForm.lbl_qty_belimlight, mainForm.lbl_qty_PRLighting,
            mainForm.lbl_qty_blackout, mainForm.lbl_qty_vision, mainForm.lbl_qty_stage}
    End Sub

    Sub initDGV()
        mainForm.dgv = {mainForm.DGV_light, mainForm.DGV_screen}
    End Sub
End Module
