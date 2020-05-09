Module declarations

    Sub initWorksheets()

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

        mainForm.tbl_Lighting_tables = {{mainForm.tbl_movHeads_belimlight, mainForm.tbl_movHeads_PRLighting, mainForm.tbl_movHeads_blackout, mainForm.tbl_movHeads_vision},
            {mainForm.tbl_strobes_belimlight, mainForm.tbl_strobes_PRLighting, mainForm.tbl_strobes_blackout, mainForm.tbl_strobes_vision},
            {mainForm.tbl_blinders_belimlight, mainForm.tbl_blinders_PRLighting, mainForm.tbl_blinders_blackout, mainForm.tbl_blinders_vision},
            {mainForm.tbl_arch_belimlight, mainForm.tbl_arch_PRLighting, mainForm.tbl_arch_blackout, mainForm.tbl_arch_vision},
            {mainForm.tbl_LED_belimlight, mainForm.tbl_LED_PRLighting, mainForm.tbl_LED_blackout, mainForm.tbl_LED_vision},
            {mainForm.tbl_smoke_belimlight, mainForm.tbl_smoke_PRLighting, mainForm.tbl_smoke_blackout, mainForm.tbl_smoke_vision},
            {mainForm.tbl_consoles_belimlight, mainForm.tbl_consoles_PRLighting, mainForm.tbl_consoles_blackout, mainForm.tbl_consoles_vision},
            {mainForm.tbl_intercom_belimlight, mainForm.tbl_intercom_PRLighting, mainForm.tbl_intercom_blackout, mainForm.tbl_intercom_vision}}

    End Sub

End Module
