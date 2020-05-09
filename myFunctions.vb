Module myFunctions

    Sub initMovHeads()
        'mainForm.tbl_Lighting_tables(0, 0) = mainForm.wsLight(0).Tables.Item("movHeads_belimlight")
        'mainForm.tbl_Lighting_tables(0, 1) = mainForm.wsLight(0).Tables.Item("movHeads_PRLighting")
        'mainForm.tbl_Lighting_tables(0, 2) = mainForm.wsLight(0).Tables.Item("movHeads_blackout")
        'mainForm.tbl_Lighting_tables(0, 3) = mainForm.wsLight(0).Tables.Item("movHeads_vision")

        mainForm.r_Light_tbl(0, 0) = mainForm.tbl_Lighting_tables(0, 0).Address.Rows
        mainForm.r_Light_tbl(0, 1) = mainForm.tbl_Lighting_tables(0, 1).Address.Rows
        mainForm.r_Light_tbl(0, 2) = mainForm.tbl_Lighting_tables(0, 2).Address.Rows
        mainForm.r_Light_tbl(0, 3) = mainForm.tbl_Lighting_tables(0, 3).Address.Rows

        mainForm.c_Light_tbl(0, 0) = mainForm.tbl_Lighting_tables(0, 0).Address.Columns
        mainForm.c_Light_tbl(0, 1) = mainForm.tbl_Lighting_tables(0, 1).Address.Columns
        mainForm.c_Light_tbl(0, 2) = mainForm.tbl_Lighting_tables(0, 2).Address.Columns
        mainForm.c_Light_tbl(0, 3) = mainForm.tbl_Lighting_tables(0, 3).Address.Columns

        mainForm.adr_Light_tbl(0, 0) = mainForm.tbl_Lighting_tables(0, 0).Address.Address
        mainForm.adr_Light_tbl(0, 1) = mainForm.tbl_Lighting_tables(0, 1).Address.Address
        mainForm.adr_Light_tbl(0, 2) = mainForm.tbl_Lighting_tables(0, 2).Address.Address
        mainForm.adr_Light_tbl(0, 3) = mainForm.tbl_Lighting_tables(0, 3).Address.Address

        mainForm.rng_Light_tbl(0, 0) = mainForm.wsMovHeads.Cells(mainForm.adr_Light_tbl(0, 0))
        mainForm.rng_Light_tbl(0, 1) = mainForm.wsMovHeads.Cells(mainForm.adr_Light_tbl(0, 1))
        mainForm.rng_Light_tbl(0, 2) = mainForm.wsMovHeads.Cells(mainForm.adr_Light_tbl(0, 2))
        mainForm.rng_Light_tbl(0, 3) = mainForm.wsMovHeads.Cells(mainForm.adr_Light_tbl(0, 3))
    End Sub

    Sub initStrobes()
        'mainForm.tbl_Lighting_tables(1, 0) = mainForm.wsLight(1).Tables.Item("strobes_belimlight")
        'mainForm.tbl_Lighting_tables(1, 1) = mainForm.wsLight(1).Tables.Item("strobes_PRLighting")
        'mainForm.tbl_Lighting_tables(1, 2) = mainForm.wsLight(1).Tables.Item("strobes_blackout")
        'mainForm.tbl_Lighting_tables(1, 3) = mainForm.wsLight(1).Tables.Item("strobes_vision")

        mainForm.r_Light_tbl(1, 0) = mainForm.tbl_Lighting_tables(1, 0).Address.Rows
        mainForm.r_Light_tbl(1, 1) = mainForm.tbl_Lighting_tables(1, 1).Address.Rows
        mainForm.r_Light_tbl(1, 2) = mainForm.tbl_Lighting_tables(1, 2).Address.Rows
        mainForm.r_Light_tbl(1, 3) = mainForm.tbl_Lighting_tables(1, 3).Address.Rows

        mainForm.c_Light_tbl(1, 0) = mainForm.tbl_Lighting_tables(1, 0).Address.Columns
        mainForm.c_Light_tbl(1, 1) = mainForm.tbl_Lighting_tables(1, 1).Address.Columns
        mainForm.c_Light_tbl(1, 2) = mainForm.tbl_Lighting_tables(1, 2).Address.Columns
        mainForm.c_Light_tbl(1, 3) = mainForm.tbl_Lighting_tables(1, 3).Address.Columns

        mainForm.adr_Light_tbl(1, 0) = mainForm.tbl_Lighting_tables(1, 0).Address.Address
        mainForm.adr_Light_tbl(1, 1) = mainForm.tbl_Lighting_tables(1, 1).Address.Address
        mainForm.adr_Light_tbl(1, 2) = mainForm.tbl_Lighting_tables(1, 2).Address.Address
        mainForm.adr_Light_tbl(1, 3) = mainForm.tbl_Lighting_tables(1, 3).Address.Address

        mainForm.rng_Light_tbl(1, 0) = mainForm.wsStrobes.Cells(mainForm.adr_Light_tbl(1, 0))
        mainForm.rng_Light_tbl(1, 1) = mainForm.wsStrobes.Cells(mainForm.adr_Light_tbl(1, 1))
        mainForm.rng_Light_tbl(1, 2) = mainForm.wsStrobes.Cells(mainForm.adr_Light_tbl(1, 2))
        mainForm.rng_Light_tbl(1, 3) = mainForm.wsStrobes.Cells(mainForm.adr_Light_tbl(1, 3))
    End Sub

    Sub initBlinders()
        'mainForm.tbl_Lighting_tables(2, 0) = mainForm.wsLight(2).Tables.Item("blinders_belimlight")
        'mainForm.tbl_Lighting_tables(2, 1) = mainForm.wsLight(2).Tables.Item("blinders_PRLighting")
        'mainForm.tbl_Lighting_tables(2, 2) = mainForm.wsLight(2).Tables.Item("blinders_blackout")
        'mainForm.tbl_Lighting_tables(2, 3) = mainForm.wsLight(2).Tables.Item("blinders_vision")

        mainForm.r_Light_tbl(2, 0) = mainForm.tbl_Lighting_tables(2, 0).Address.Rows
        mainForm.r_Light_tbl(2, 1) = mainForm.tbl_Lighting_tables(2, 1).Address.Rows
        mainForm.r_Light_tbl(2, 2) = mainForm.tbl_Lighting_tables(2, 2).Address.Rows
        mainForm.r_Light_tbl(2, 3) = mainForm.tbl_Lighting_tables(2, 3).Address.Rows

        mainForm.c_Light_tbl(2, 0) = mainForm.tbl_Lighting_tables(2, 0).Address.Columns
        mainForm.c_Light_tbl(2, 1) = mainForm.tbl_Lighting_tables(2, 1).Address.Columns
        mainForm.c_Light_tbl(2, 2) = mainForm.tbl_Lighting_tables(2, 2).Address.Columns
        mainForm.c_Light_tbl(2, 3) = mainForm.tbl_Lighting_tables(2, 3).Address.Columns

        mainForm.adr_Light_tbl(2, 0) = mainForm.tbl_Lighting_tables(2, 0).Address.Address
        mainForm.adr_Light_tbl(2, 1) = mainForm.tbl_Lighting_tables(2, 1).Address.Address
        mainForm.adr_Light_tbl(2, 2) = mainForm.tbl_Lighting_tables(2, 2).Address.Address
        mainForm.adr_Light_tbl(2, 3) = mainForm.tbl_Lighting_tables(2, 3).Address.Address

        mainForm.rng_Light_tbl(2, 0) = mainForm.wsBlinders.Cells(mainForm.adr_Light_tbl(2, 0))
        mainForm.rng_Light_tbl(2, 1) = mainForm.wsBlinders.Cells(mainForm.adr_Light_tbl(2, 1))
        mainForm.rng_Light_tbl(2, 2) = mainForm.wsBlinders.Cells(mainForm.adr_Light_tbl(2, 2))
        mainForm.rng_Light_tbl(2, 3) = mainForm.wsBlinders.Cells(mainForm.adr_Light_tbl(2, 3))
    End Sub

    Sub initArch()
        'mainForm.tbl_Lighting_tables(3, 0) = mainForm.wsLight(3).Tables.Item("arch_belimlight")
        'mainForm.tbl_Lighting_tables(3, 1) = mainForm.wsLight(3).Tables.Item("arch_PRLighting")
        'mainForm.tbl_Lighting_tables(3, 2) = mainForm.wsLight(3).Tables.Item("arch_blackout")
        'mainForm.tbl_Lighting_tables(3, 3) = mainForm.wsLight(3).Tables.Item("arch_vision")

        mainForm.r_Light_tbl(3, 0) = mainForm.tbl_Lighting_tables(3, 0).Address.Rows
        mainForm.r_Light_tbl(3, 1) = mainForm.tbl_Lighting_tables(3, 1).Address.Rows
        mainForm.r_Light_tbl(3, 2) = mainForm.tbl_Lighting_tables(3, 2).Address.Rows
        mainForm.r_Light_tbl(3, 3) = mainForm.tbl_Lighting_tables(3, 3).Address.Rows

        mainForm.c_Light_tbl(3, 0) = mainForm.tbl_Lighting_tables(3, 0).Address.Columns
        mainForm.c_Light_tbl(3, 1) = mainForm.tbl_Lighting_tables(3, 1).Address.Columns
        mainForm.c_Light_tbl(3, 2) = mainForm.tbl_Lighting_tables(3, 2).Address.Columns
        mainForm.c_Light_tbl(3, 3) = mainForm.tbl_Lighting_tables(3, 3).Address.Columns

        mainForm.adr_Light_tbl(3, 0) = mainForm.tbl_Lighting_tables(3, 0).Address.Address
        mainForm.adr_Light_tbl(3, 1) = mainForm.tbl_Lighting_tables(3, 1).Address.Address
        mainForm.adr_Light_tbl(3, 2) = mainForm.tbl_Lighting_tables(3, 2).Address.Address
        mainForm.adr_Light_tbl(3, 3) = mainForm.tbl_Lighting_tables(3, 3).Address.Address

        mainForm.rng_Light_tbl(3, 0) = mainForm.wsArch.Cells(mainForm.adr_Light_tbl(3, 0))
        mainForm.rng_Light_tbl(3, 1) = mainForm.wsArch.Cells(mainForm.adr_Light_tbl(3, 1))
        mainForm.rng_Light_tbl(3, 2) = mainForm.wsArch.Cells(mainForm.adr_Light_tbl(3, 2))
        mainForm.rng_Light_tbl(3, 3) = mainForm.wsArch.Cells(mainForm.adr_Light_tbl(3, 3))
    End Sub

    Sub initLED()
        'mainForm.tbl_Lighting_tables(4, 0) = mainForm.wsLight(4).Tables.Item("LED_belimlight")
        'mainForm.tbl_Lighting_tables(4, 1) = mainForm.wsLight(4).Tables.Item("LED_PRLighting")
        'mainForm.tbl_Lighting_tables(4, 2) = mainForm.wsLight(4).Tables.Item("LED_blackout")
        'mainForm.tbl_Lighting_tables(4, 3) = mainForm.wsLight(4).Tables.Item("LED_vision")

        mainForm.r_Light_tbl(4, 0) = mainForm.tbl_Lighting_tables(4, 0).Address.Rows
        mainForm.r_Light_tbl(4, 1) = mainForm.tbl_Lighting_tables(4, 1).Address.Rows
        mainForm.r_Light_tbl(4, 2) = mainForm.tbl_Lighting_tables(4, 2).Address.Rows
        mainForm.r_Light_tbl(4, 3) = mainForm.tbl_Lighting_tables(4, 3).Address.Rows

        mainForm.c_Light_tbl(4, 0) = mainForm.tbl_Lighting_tables(4, 0).Address.Columns
        mainForm.c_Light_tbl(4, 1) = mainForm.tbl_Lighting_tables(4, 1).Address.Columns
        mainForm.c_Light_tbl(4, 2) = mainForm.tbl_Lighting_tables(4, 2).Address.Columns
        mainForm.c_Light_tbl(4, 3) = mainForm.tbl_Lighting_tables(4, 3).Address.Columns

        mainForm.adr_Light_tbl(4, 0) = mainForm.tbl_Lighting_tables(4, 0).Address.Address
        mainForm.adr_Light_tbl(4, 1) = mainForm.tbl_Lighting_tables(4, 1).Address.Address
        mainForm.adr_Light_tbl(4, 2) = mainForm.tbl_Lighting_tables(4, 2).Address.Address
        mainForm.adr_Light_tbl(4, 3) = mainForm.tbl_Lighting_tables(4, 3).Address.Address

        mainForm.rng_Light_tbl(4, 0) = mainForm.wsLED.Cells(mainForm.adr_Light_tbl(4, 0))
        mainForm.rng_Light_tbl(4, 1) = mainForm.wsLED.Cells(mainForm.adr_Light_tbl(4, 1))
        mainForm.rng_Light_tbl(4, 2) = mainForm.wsLED.Cells(mainForm.adr_Light_tbl(4, 2))
        mainForm.rng_Light_tbl(4, 3) = mainForm.wsLED.Cells(mainForm.adr_Light_tbl(4, 3))
    End Sub

    Sub initSmoke()
        'mainForm.tbl_Lighting_tables(5, 0) = mainForm.wsLight(5).Tables.Item("smoke_belimlight")
        'mainForm.tbl_Lighting_tables(5, 1) = mainForm.wsLight(5).Tables.Item("smoke_PRLighting")
        'mainForm.tbl_Lighting_tables(5, 2) = mainForm.wsLight(5).Tables.Item("smoke_blackout")
        'mainForm.tbl_Lighting_tables(5, 3) = mainForm.wsLight(5).Tables.Item("smoke_vision")

        mainForm.r_Light_tbl(5, 0) = mainForm.tbl_Lighting_tables(5, 0).Address.Rows
        mainForm.r_Light_tbl(5, 1) = mainForm.tbl_Lighting_tables(5, 1).Address.Rows
        mainForm.r_Light_tbl(5, 2) = mainForm.tbl_Lighting_tables(5, 2).Address.Rows
        mainForm.r_Light_tbl(5, 3) = mainForm.tbl_Lighting_tables(5, 3).Address.Rows

        mainForm.c_Light_tbl(5, 0) = mainForm.tbl_Lighting_tables(5, 0).Address.Columns
        mainForm.c_Light_tbl(5, 1) = mainForm.tbl_Lighting_tables(5, 1).Address.Columns
        mainForm.c_Light_tbl(5, 2) = mainForm.tbl_Lighting_tables(5, 2).Address.Columns
        mainForm.c_Light_tbl(5, 3) = mainForm.tbl_Lighting_tables(5, 3).Address.Columns

        mainForm.adr_Light_tbl(5, 0) = mainForm.tbl_Lighting_tables(5, 0).Address.Address
        mainForm.adr_Light_tbl(5, 1) = mainForm.tbl_Lighting_tables(5, 1).Address.Address
        mainForm.adr_Light_tbl(5, 2) = mainForm.tbl_Lighting_tables(5, 2).Address.Address
        mainForm.adr_Light_tbl(5, 3) = mainForm.tbl_Lighting_tables(5, 3).Address.Address

        mainForm.rng_Light_tbl(5, 0) = mainForm.wsSmoke.Cells(mainForm.adr_Light_tbl(5, 0))
        mainForm.rng_Light_tbl(5, 1) = mainForm.wsSmoke.Cells(mainForm.adr_Light_tbl(5, 1))
        mainForm.rng_Light_tbl(5, 2) = mainForm.wsSmoke.Cells(mainForm.adr_Light_tbl(5, 2))
        mainForm.rng_Light_tbl(5, 3) = mainForm.wsSmoke.Cells(mainForm.adr_Light_tbl(5, 3))
    End Sub

    Sub initConsoles()
        'mainForm.tbl_Lighting_tables(6, 0) = mainForm.wsLight(6).Tables.Item("consoles_belimlight")
        'mainForm.tbl_Lighting_tables(6, 1) = mainForm.wsLight(6).Tables.Item("consoles_PRLighting")
        'mainForm.tbl_Lighting_tables(6, 2) = mainForm.wsLight(6).Tables.Item("consoles_blackout")
        'mainForm.tbl_Lighting_tables(6, 3) = mainForm.wsLight(6).Tables.Item("consoles_vision")

        mainForm.r_Light_tbl(6, 0) = mainForm.tbl_Lighting_tables(6, 0).Address.Rows
        mainForm.r_Light_tbl(6, 1) = mainForm.tbl_Lighting_tables(6, 1).Address.Rows
        mainForm.r_Light_tbl(6, 2) = mainForm.tbl_Lighting_tables(6, 2).Address.Rows
        mainForm.r_Light_tbl(6, 3) = mainForm.tbl_Lighting_tables(6, 3).Address.Rows

        mainForm.c_Light_tbl(6, 0) = mainForm.tbl_Lighting_tables(6, 0).Address.Columns
        mainForm.c_Light_tbl(6, 1) = mainForm.tbl_Lighting_tables(6, 1).Address.Columns
        mainForm.c_Light_tbl(6, 2) = mainForm.tbl_Lighting_tables(6, 2).Address.Columns
        mainForm.c_Light_tbl(6, 3) = mainForm.tbl_Lighting_tables(6, 3).Address.Columns

        mainForm.adr_Light_tbl(6, 0) = mainForm.tbl_Lighting_tables(6, 0).Address.Address
        mainForm.adr_Light_tbl(6, 1) = mainForm.tbl_Lighting_tables(6, 1).Address.Address
        mainForm.adr_Light_tbl(6, 2) = mainForm.tbl_Lighting_tables(6, 2).Address.Address
        mainForm.adr_Light_tbl(6, 3) = mainForm.tbl_Lighting_tables(6, 3).Address.Address

        mainForm.rng_Light_tbl(6, 0) = mainForm.wsConsoles.Cells(mainForm.adr_Light_tbl(6, 0))
        mainForm.rng_Light_tbl(6, 1) = mainForm.wsConsoles.Cells(mainForm.adr_Light_tbl(6, 1))
        mainForm.rng_Light_tbl(6, 2) = mainForm.wsConsoles.Cells(mainForm.adr_Light_tbl(6, 2))
        mainForm.rng_Light_tbl(6, 3) = mainForm.wsConsoles.Cells(mainForm.adr_Light_tbl(6, 3))
    End Sub

    Sub initIntercom()
        'mainForm.tbl_Lighting_tables(7, 0) = mainForm.wsLight(7).Tables.Item("intercom_belimlight")
        'mainForm.tbl_Lighting_tables(7, 1) = mainForm.wsLight(7).Tables.Item("intercom_PRLighting")
        'mainForm.tbl_Lighting_tables(7, 2) = mainForm.wsLight(7).Tables.Item("intercom_blackout")
        'mainForm.tbl_Lighting_tables(7, 3) = mainForm.wsLight(7).Tables.Item("intercom_vision")

        mainForm.r_Light_tbl(7, 0) = mainForm.tbl_Lighting_tables(7, 0).Address.Rows
        mainForm.r_Light_tbl(7, 1) = mainForm.tbl_Lighting_tables(7, 1).Address.Rows
        mainForm.r_Light_tbl(7, 2) = mainForm.tbl_Lighting_tables(7, 2).Address.Rows
        mainForm.r_Light_tbl(7, 3) = mainForm.tbl_Lighting_tables(7, 3).Address.Rows

        mainForm.c_Light_tbl(7, 0) = mainForm.tbl_Lighting_tables(7, 0).Address.Columns
        mainForm.c_Light_tbl(7, 1) = mainForm.tbl_Lighting_tables(7, 1).Address.Columns
        mainForm.c_Light_tbl(7, 2) = mainForm.tbl_Lighting_tables(7, 2).Address.Columns
        mainForm.c_Light_tbl(7, 3) = mainForm.tbl_Lighting_tables(7, 3).Address.Columns

        mainForm.adr_Light_tbl(7, 0) = mainForm.tbl_Lighting_tables(7, 0).Address.Address
        mainForm.adr_Light_tbl(7, 1) = mainForm.tbl_Lighting_tables(7, 1).Address.Address
        mainForm.adr_Light_tbl(7, 2) = mainForm.tbl_Lighting_tables(7, 2).Address.Address
        mainForm.adr_Light_tbl(7, 3) = mainForm.tbl_Lighting_tables(7, 3).Address.Address

        mainForm.rng_Light_tbl(7, 0) = mainForm.wsIntercom.Cells(mainForm.adr_Light_tbl(7, 0))
        mainForm.rng_Light_tbl(7, 1) = mainForm.wsIntercom.Cells(mainForm.adr_Light_tbl(7, 1))
        mainForm.rng_Light_tbl(7, 2) = mainForm.wsIntercom.Cells(mainForm.adr_Light_tbl(7, 2))
        mainForm.rng_Light_tbl(7, 3) = mainForm.wsIntercom.Cells(mainForm.adr_Light_tbl(7, 3))
    End Sub


End Module
