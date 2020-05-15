Module declarations

    Sub initLightWorksheets()

        mainForm.wsMovHeads = mainForm.obj_excel.Workbook.Worksheets(0)
        mainForm.wsStrobes = mainForm.obj_excel.Workbook.Worksheets(1)
        mainForm.wsBlinders = mainForm.obj_excel.Workbook.Worksheets(2)
        mainForm.wsArch = mainForm.obj_excel.Workbook.Worksheets(3)
        mainForm.wsLED = mainForm.obj_excel.Workbook.Worksheets(4)
        mainForm.wsSmoke = mainForm.obj_excel.Workbook.Worksheets(5)
        mainForm.wsConsoles = mainForm.obj_excel.Workbook.Worksheets(6)
        mainForm.wsIntercom = mainForm.obj_excel.Workbook.Worksheets(7)

        mainForm.wsLight = {mainForm.wsMovHeads,
                            mainForm.wsStrobes,
                            mainForm.wsBlinders,
                            mainForm.wsArch,
                            mainForm.wsLED,
                            mainForm.wsSmoke,
                            mainForm.wsConsoles,
                            mainForm.wsIntercom}



    End Sub

    Sub initLightTables()

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

        mainForm.tbl_movHeads = mainForm.wsLight(0).Tables.Item("movHeads_tbl")
        mainForm.tbl_strobes = mainForm.wsLight(1).Tables.Item("strobes_tbl")
        mainForm.tbl_blinders = mainForm.wsLight(2).Tables.Item("blinders_tbl")
        mainForm.tbl_arch = mainForm.wsLight(3).Tables.Item("arch_tbl")
        mainForm.tbl_LED = mainForm.wsLight(4).Tables.Item("LED_tbl")
        mainForm.tbl_smoke = mainForm.wsLight(5).Tables.Item("smoke_tbl")
        mainForm.tbl_consoles = mainForm.wsLight(6).Tables.Item("consoles_tbl")
        mainForm.tbl_intercom = mainForm.wsLight(7).Tables.Item("intercom_tbl")

        mainForm.tbl_Lighting_sumTables = {mainForm.tbl_movHeads, mainForm.tbl_strobes, mainForm.tbl_blinders,
            mainForm.tbl_arch, mainForm.tbl_LED, mainForm.tbl_smoke, mainForm.tbl_consoles, mainForm.tbl_intercom}

        'Console.WriteLine(mainForm.tbl_Lighting_sumTables(0).Name)

        mainForm.lblSumQty = {mainForm.lbl_qty_belimlight, mainForm.lbl_qty_PRLighting,
            mainForm.lbl_qty_blackout, mainForm.lbl_qty_vision, mainForm.lbl_qty_stage}


        mainForm.dt_Lighting = {
            {mainForm.dt_movHeads_belimlight, mainForm.dt_movHeads_PRLighting,
            mainForm.dt_movHeads_blackout, mainForm.dt_movHeads_vision, mainForm.dt_movHeads_stage},
            {mainForm.dt_strobes_belimlight, mainForm.dt_strobes_PRLighting,
            mainForm.dt_strobes_blackout, mainForm.dt_strobes_vision, mainForm.dt_strobes_stage},
            {mainForm.dt_blinders_belimlight, mainForm.dt_blinders_PRLighting,
            mainForm.dt_blinders_blackout, mainForm.dt_blinders_vision, mainForm.dt_blinders_stage},
            {mainForm.dt_arch_belimlight, mainForm.dt_arch_PRLighting,
            mainForm.dt_arch_blackout, mainForm.dt_arch_vision, mainForm.dt_arch_stage},
            {mainForm.dt_LED_belimlight, mainForm.dt_LED_PRLighting,
            mainForm.dt_LED_blackout, mainForm.dt_LED_vision, mainForm.dt_LED_stage},
            {mainForm.dt_smoke_belimlight, mainForm.dt_smoke_PRLighting,
            mainForm.dt_smoke_blackout, mainForm.dt_smoke_vision, mainForm.dt_smoke_stage},
            {mainForm.dt_consoles_belimlight, mainForm.dt_consoles_PRLighting,
            mainForm.dt_consoles_blackout, mainForm.dt_consoles_vision, mainForm.dt_consoles_stage},
            {mainForm.dt_intercom_belimlight, mainForm.dt_intercom_PRLighting,
            mainForm.dt_intercom_blackout, mainForm.dt_intercom_vision, mainForm.dt_intercom_stage}
            }


        mainForm.dt_sumLighting = {mainForm.dt_movHeads, mainForm.dt_strobes, mainForm.dt_blinders,
            mainForm.dt_arch, mainForm.dt_LED, mainForm.dt_smoke, mainForm.dt_consoles, mainForm.dt_intercom}



        Dim i, j As Integer

        For i = 0 To 7

            For j = 0 To 4

                mainForm.r_Light_tbl(i, j) = mainForm.tbl_Lighting_tables(i, j).Address.Rows
                mainForm.c_Light_tbl(i, j) = mainForm.tbl_Lighting_tables(i, j).Address.Columns
                mainForm.adr_Light_tbl(i, j) = mainForm.tbl_Lighting_tables(i, j).Address.Address
                mainForm.rng_Light_tbl(i, j) = mainForm.wsLight(i).Cells(mainForm.adr_Light_tbl(i, j))

            Next j
            'Console.WriteLine(mainForm.tbl_Lighting_sumTables(0).Name)
            mainForm.r_Light_sumTbl(i) = mainForm.tbl_Lighting_sumTables(i).Address.Rows
            mainForm.c_Light_sumTbl(i) = mainForm.tbl_Lighting_sumTables(i).Address.Columns
            mainForm.adr_Light_sumTbl(i) = mainForm.tbl_Lighting_sumTables(i).Address.Address
            mainForm.rng_Light_sumTbl(i) = mainForm.wsLight(i).Cells(mainForm.adr_Light_sumTbl(i))

        Next i

    End Sub

End Module
