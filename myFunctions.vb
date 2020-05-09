Module myFunctions

    Sub initMovHeads()
        mainForm.tbl_Lighting_tables(0, 0) = mainForm.wsLight(0).Tables.Item("movHeads_belimlight")
        mainForm.tbl_Lighting_tables(0, 1) = mainForm.wsLight(0).Tables.Item("movHeads_PRLighting")
        mainForm.tbl_Lighting_tables(0, 2) = mainForm.wsLight(0).Tables.Item("movHeads_blackout")
        mainForm.tbl_Lighting_tables(0, 3) = mainForm.wsLight(0).Tables.Item("movHeads_vision")

        mainForm.r_movHeads_belimlight = mainForm.tbl_Lighting_tables(0, 0).Address.Rows
        mainForm.r_movHeads_PRLighting = mainForm.tbl_Lighting_tables(0, 1).Address.Rows
        mainForm.r_movHeads_blackout = mainForm.tbl_Lighting_tables(0, 2).Address.Rows
        mainForm.r_movHeads_vision = mainForm.tbl_Lighting_tables(0, 3).Address.Rows

        mainForm.c_movHeads_belimlight = mainForm.tbl_Lighting_tables(0, 0).Address.Columns
        mainForm.c_movHeads_PRLighting = mainForm.tbl_Lighting_tables(0, 1).Address.Columns
        mainForm.c_movHeads_blackout = mainForm.tbl_Lighting_tables(0, 2).Address.Columns
        mainForm.c_movHeads_vision = mainForm.tbl_Lighting_tables(0, 3).Address.Columns

        mainForm.adr_movHeads_belimlight = mainForm.tbl_Lighting_tables(0, 0).Address.Address
        mainForm.adr_movHeads_PRLighting = mainForm.tbl_Lighting_tables(0, 1).Address.Address
        mainForm.adr_movHeads_blackout = mainForm.tbl_Lighting_tables(0, 2).Address.Address
        mainForm.adr_movHeads_vision = mainForm.tbl_Lighting_tables(0, 3).Address.Address

        mainForm.rng_movHeads_belimlight = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_belimlight)
        mainForm.rng_movHeads_PRLighting = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_PRLighting)
        mainForm.rng_movHeads_blackout = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_blackout)
        mainForm.rng_movHeads_vision = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_vision)
    End Sub

    Sub initStrobes()
        mainForm.tbl_Lighting_tables(1, 0) = mainForm.wsLight(1).Tables.Item("strobes_belimlight")
        mainForm.tbl_Lighting_tables(1, 1) = mainForm.wsLight(1).Tables.Item("strobes_PRLighting")
        mainForm.tbl_Lighting_tables(1, 2) = mainForm.wsLight(1).Tables.Item("strobes_blackout")
        mainForm.tbl_Lighting_tables(1, 3) = mainForm.wsLight(1).Tables.Item("strobes_vision")

        mainForm.r_strobes_belimlight = mainForm.tbl_Lighting_tables(1, 0).Address.Rows
        mainForm.r_strobes_PRLighting = mainForm.tbl_Lighting_tables(1, 1).Address.Rows
        mainForm.r_strobes_blackout = mainForm.tbl_Lighting_tables(1, 2).Address.Rows
        mainForm.r_strobes_vision = mainForm.tbl_Lighting_tables(1, 3).Address.Rows

        mainForm.c_strobes_belimlight = mainForm.tbl_Lighting_tables(1, 0).Address.Columns
        mainForm.c_strobes_PRLighting = mainForm.tbl_Lighting_tables(1, 1).Address.Columns
        mainForm.c_strobes_blackout = mainForm.tbl_Lighting_tables(1, 2).Address.Columns
        mainForm.c_strobes_vision = mainForm.tbl_Lighting_tables(1, 3).Address.Columns

        mainForm.adr_strobes_belimlight = mainForm.tbl_Lighting_tables(1, 0).Address.Address
        mainForm.adr_strobes_PRLighting = mainForm.tbl_Lighting_tables(1, 1).Address.Address
        mainForm.adr_strobes_blackout = mainForm.tbl_Lighting_tables(1, 2).Address.Address
        mainForm.adr_strobes_vision = mainForm.tbl_Lighting_tables(1, 3).Address.Address

        mainForm.rng_strobes_belimlight = mainForm.wsStrobes.Cells(mainForm.adr_strobes_belimlight)
        mainForm.rng_strobes_PRLighting = mainForm.wsStrobes.Cells(mainForm.adr_strobes_PRLighting)
        mainForm.rng_strobes_blackout = mainForm.wsStrobes.Cells(mainForm.adr_strobes_blackout)
        mainForm.rng_strobes_vision = mainForm.wsStrobes.Cells(mainForm.adr_strobes_vision)
    End Sub

    Sub initBlinders()
        mainForm.tbl_Lighting_tables(2, 0) = mainForm.wsLight(2).Tables.Item("blinders_belimlight")
        mainForm.tbl_Lighting_tables(2, 1) = mainForm.wsLight(2).Tables.Item("blinders_PRLighting")
        mainForm.tbl_Lighting_tables(2, 2) = mainForm.wsLight(2).Tables.Item("blinders_blackout")
        mainForm.tbl_Lighting_tables(2, 3) = mainForm.wsLight(2).Tables.Item("blinders_vision")

        mainForm.r_blinders_belimlight = mainForm.tbl_Lighting_tables(2, 0).Address.Rows
        mainForm.r_blinders_PRLighting = mainForm.tbl_Lighting_tables(2, 1).Address.Rows
        mainForm.r_blinders_blackout = mainForm.tbl_Lighting_tables(2, 2).Address.Rows
        mainForm.r_blinders_vision = mainForm.tbl_Lighting_tables(2, 3).Address.Rows

        mainForm.c_blinders_belimlight = mainForm.tbl_Lighting_tables(2, 0).Address.Columns
        mainForm.c_blinders_PRLighting = mainForm.tbl_Lighting_tables(2, 1).Address.Columns
        mainForm.c_blinders_blackout = mainForm.tbl_Lighting_tables(2, 2).Address.Columns
        mainForm.c_blinders_vision = mainForm.tbl_Lighting_tables(2, 3).Address.Columns

        mainForm.adr_blinders_belimlight = mainForm.tbl_Lighting_tables(2, 0).Address.Address
        mainForm.adr_blinders_PRLighting = mainForm.tbl_Lighting_tables(2, 1).Address.Address
        mainForm.adr_blinders_blackout = mainForm.tbl_Lighting_tables(2, 2).Address.Address
        mainForm.adr_blinders_vision = mainForm.tbl_Lighting_tables(2, 3).Address.Address

        mainForm.rng_blinders_belimlight = mainForm.wsblinders.Cells(mainForm.adr_blinders_belimlight)
        mainForm.rng_blinders_PRLighting = mainForm.wsblinders.Cells(mainForm.adr_blinders_PRLighting)
        mainForm.rng_blinders_blackout = mainForm.wsblinders.Cells(mainForm.adr_blinders_blackout)
        mainForm.rng_blinders_vision = mainForm.wsblinders.Cells(mainForm.adr_blinders_vision)
    End Sub

    Sub initArch()
        mainForm.tbl_Lighting_tables(3, 0) = mainForm.wsLight(3).Tables.Item("arch_belimlight")
        mainForm.tbl_Lighting_tables(3, 1) = mainForm.wsLight(3).Tables.Item("arch_PRLighting")
        mainForm.tbl_Lighting_tables(3, 2) = mainForm.wsLight(3).Tables.Item("arch_blackout")
        mainForm.tbl_Lighting_tables(3, 3) = mainForm.wsLight(3).Tables.Item("arch_vision")

        mainForm.r_arch_belimlight = mainForm.tbl_Lighting_tables(3, 0).Address.Rows
        mainForm.r_arch_PRLighting = mainForm.tbl_Lighting_tables(3, 1).Address.Rows
        mainForm.r_arch_blackout = mainForm.tbl_Lighting_tables(3, 2).Address.Rows
        mainForm.r_arch_vision = mainForm.tbl_Lighting_tables(3, 3).Address.Rows

        mainForm.c_arch_belimlight = mainForm.tbl_Lighting_tables(3, 0).Address.Columns
        mainForm.c_arch_PRLighting = mainForm.tbl_Lighting_tables(3, 1).Address.Columns
        mainForm.c_arch_blackout = mainForm.tbl_Lighting_tables(3, 2).Address.Columns
        mainForm.c_arch_vision = mainForm.tbl_Lighting_tables(3, 3).Address.Columns

        mainForm.adr_arch_belimlight = mainForm.tbl_Lighting_tables(3, 0).Address.Address
        mainForm.adr_arch_PRLighting = mainForm.tbl_Lighting_tables(3, 1).Address.Address
        mainForm.adr_arch_blackout = mainForm.tbl_Lighting_tables(3, 2).Address.Address
        mainForm.adr_arch_vision = mainForm.tbl_Lighting_tables(3, 3).Address.Address

        mainForm.rng_arch_belimlight = mainForm.wsArch.Cells(mainForm.adr_arch_belimlight)
        mainForm.rng_arch_PRLighting = mainForm.wsArch.Cells(mainForm.adr_arch_PRLighting)
        mainForm.rng_arch_blackout = mainForm.wsArch.Cells(mainForm.adr_arch_blackout)
        mainForm.rng_arch_vision = mainForm.wsArch.Cells(mainForm.adr_arch_vision)
    End Sub

    Sub initLED()
        mainForm.tbl_Lighting_tables(4, 0) = mainForm.wsLight(4).Tables.Item("LED_belimlight")
        mainForm.tbl_Lighting_tables(4, 1) = mainForm.wsLight(4).Tables.Item("LED_PRLighting")
        mainForm.tbl_Lighting_tables(4, 2) = mainForm.wsLight(4).Tables.Item("LED_blackout")
        mainForm.tbl_Lighting_tables(4, 3) = mainForm.wsLight(4).Tables.Item("LED_vision")

        mainForm.r_LED_belimlight = mainForm.tbl_Lighting_tables(4, 0).Address.Rows
        mainForm.r_LED_PRLighting = mainForm.tbl_Lighting_tables(4, 1).Address.Rows
        mainForm.r_LED_blackout = mainForm.tbl_Lighting_tables(4, 2).Address.Rows
        mainForm.r_LED_vision = mainForm.tbl_Lighting_tables(4, 3).Address.Rows

        mainForm.c_LED_belimlight = mainForm.tbl_Lighting_tables(4, 0).Address.Columns
        mainForm.c_LED_PRLighting = mainForm.tbl_Lighting_tables(4, 1).Address.Columns
        mainForm.c_LED_blackout = mainForm.tbl_Lighting_tables(4, 2).Address.Columns
        mainForm.c_LED_vision = mainForm.tbl_Lighting_tables(4, 3).Address.Columns

        mainForm.adr_LED_belimlight = mainForm.tbl_Lighting_tables(4, 0).Address.Address
        mainForm.adr_LED_PRLighting = mainForm.tbl_Lighting_tables(4, 1).Address.Address
        mainForm.adr_LED_blackout = mainForm.tbl_Lighting_tables(4, 2).Address.Address
        mainForm.adr_LED_vision = mainForm.tbl_Lighting_tables(4, 3).Address.Address

        mainForm.rng_LED_belimlight = mainForm.wsLED.Cells(mainForm.adr_LED_belimlight)
        mainForm.rng_LED_PRLighting = mainForm.wsLED.Cells(mainForm.adr_LED_PRLighting)
        mainForm.rng_LED_blackout = mainForm.wsLED.Cells(mainForm.adr_LED_blackout)
        mainForm.rng_LED_vision = mainForm.wsLED.Cells(mainForm.adr_LED_vision)
    End Sub

    Sub initSmoke()
        mainForm.tbl_Lighting_tables(5, 0) = mainForm.wsLight(5).Tables.Item("smoke_belimlight")
        mainForm.tbl_Lighting_tables(5, 1) = mainForm.wsLight(5).Tables.Item("smoke_PRLighting")
        mainForm.tbl_Lighting_tables(5, 2) = mainForm.wsLight(5).Tables.Item("smoke_blackout")
        mainForm.tbl_Lighting_tables(5, 3) = mainForm.wsLight(5).Tables.Item("smoke_vision")

        mainForm.r_smoke_belimlight = mainForm.tbl_Lighting_tables(5, 0).Address.Rows
        mainForm.r_smoke_PRLighting = mainForm.tbl_Lighting_tables(5, 1).Address.Rows
        mainForm.r_smoke_blackout = mainForm.tbl_Lighting_tables(5, 2).Address.Rows
        mainForm.r_smoke_vision = mainForm.tbl_Lighting_tables(5, 3).Address.Rows

        mainForm.c_smoke_belimlight = mainForm.tbl_Lighting_tables(5, 0).Address.Columns
        mainForm.c_smoke_PRLighting = mainForm.tbl_Lighting_tables(5, 1).Address.Columns
        mainForm.c_smoke_blackout = mainForm.tbl_Lighting_tables(5, 2).Address.Columns
        mainForm.c_smoke_vision = mainForm.tbl_Lighting_tables(5, 3).Address.Columns

        mainForm.adr_smoke_belimlight = mainForm.tbl_Lighting_tables(5, 0).Address.Address
        mainForm.adr_smoke_PRLighting = mainForm.tbl_Lighting_tables(5, 1).Address.Address
        mainForm.adr_smoke_blackout = mainForm.tbl_Lighting_tables(5, 2).Address.Address
        mainForm.adr_smoke_vision = mainForm.tbl_Lighting_tables(5, 3).Address.Address

        mainForm.rng_smoke_belimlight = mainForm.wsSmoke.Cells(mainForm.adr_smoke_belimlight)
        mainForm.rng_smoke_PRLighting = mainForm.wsSmoke.Cells(mainForm.adr_smoke_PRLighting)
        mainForm.rng_smoke_blackout = mainForm.wsSmoke.Cells(mainForm.adr_smoke_blackout)
        mainForm.rng_smoke_vision = mainForm.wsSmoke.Cells(mainForm.adr_smoke_vision)
    End Sub

    Sub initConsoles()
        mainForm.tbl_Lighting_tables(6, 0) = mainForm.wsLight(6).Tables.Item("consoles_belimlight")
        mainForm.tbl_Lighting_tables(6, 1) = mainForm.wsLight(6).Tables.Item("consoles_PRLighting")
        mainForm.tbl_Lighting_tables(6, 2) = mainForm.wsLight(6).Tables.Item("consoles_blackout")
        mainForm.tbl_Lighting_tables(6, 3) = mainForm.wsLight(6).Tables.Item("consoles_vision")

        mainForm.r_consoles_belimlight = mainForm.tbl_Lighting_tables(6, 0).Address.Rows
        mainForm.r_consoles_PRLighting = mainForm.tbl_Lighting_tables(6, 1).Address.Rows
        mainForm.r_consoles_blackout = mainForm.tbl_Lighting_tables(6, 2).Address.Rows
        mainForm.r_consoles_vision = mainForm.tbl_Lighting_tables(6, 3).Address.Rows

        mainForm.c_consoles_belimlight = mainForm.tbl_Lighting_tables(6, 0).Address.Columns
        mainForm.c_consoles_PRLighting = mainForm.tbl_Lighting_tables(6, 1).Address.Columns
        mainForm.c_consoles_blackout = mainForm.tbl_Lighting_tables(6, 2).Address.Columns
        mainForm.c_consoles_vision = mainForm.tbl_Lighting_tables(6, 3).Address.Columns

        mainForm.adr_consoles_belimlight = mainForm.tbl_Lighting_tables(6, 0).Address.Address
        mainForm.adr_consoles_PRLighting = mainForm.tbl_Lighting_tables(6, 1).Address.Address
        mainForm.adr_consoles_blackout = mainForm.tbl_Lighting_tables(6, 2).Address.Address
        mainForm.adr_consoles_vision = mainForm.tbl_Lighting_tables(6, 3).Address.Address

        mainForm.rng_consoles_belimlight = mainForm.wsConsoles.Cells(mainForm.adr_consoles_belimlight)
        mainForm.rng_consoles_PRLighting = mainForm.wsConsoles.Cells(mainForm.adr_consoles_PRLighting)
        mainForm.rng_consoles_blackout = mainForm.wsConsoles.Cells(mainForm.adr_consoles_blackout)
        mainForm.rng_consoles_vision = mainForm.wsConsoles.Cells(mainForm.adr_consoles_vision)
    End Sub

    Sub initIntercom()
        mainForm.tbl_Lighting_tables(7, 0) = mainForm.wsLight(7).Tables.Item("intercom_belimlight")
        mainForm.tbl_Lighting_tables(7, 1) = mainForm.wsLight(7).Tables.Item("intercom_PRLighting")
        mainForm.tbl_Lighting_tables(7, 2) = mainForm.wsLight(7).Tables.Item("intercom_blackout")
        mainForm.tbl_Lighting_tables(7, 3) = mainForm.wsLight(7).Tables.Item("intercom_vision")

        mainForm.r_intercom_belimlight = mainForm.tbl_Lighting_tables(7, 0).Address.Rows
        mainForm.r_intercom_PRLighting = mainForm.tbl_Lighting_tables(7, 1).Address.Rows
        mainForm.r_intercom_blackout = mainForm.tbl_Lighting_tables(7, 2).Address.Rows
        mainForm.r_intercom_vision = mainForm.tbl_Lighting_tables(7, 3).Address.Rows

        mainForm.c_intercom_belimlight = mainForm.tbl_Lighting_tables(7, 0).Address.Columns
        mainForm.c_intercom_PRLighting = mainForm.tbl_Lighting_tables(7, 1).Address.Columns
        mainForm.c_intercom_blackout = mainForm.tbl_Lighting_tables(7, 2).Address.Columns
        mainForm.c_intercom_vision = mainForm.tbl_Lighting_tables(7, 3).Address.Columns

        mainForm.adr_intercom_belimlight = mainForm.tbl_Lighting_tables(7, 0).Address.Address
        mainForm.adr_intercom_PRLighting = mainForm.tbl_Lighting_tables(7, 1).Address.Address
        mainForm.adr_intercom_blackout = mainForm.tbl_Lighting_tables(7, 2).Address.Address
        mainForm.adr_intercom_vision = mainForm.tbl_Lighting_tables(7, 3).Address.Address

        mainForm.rng_intercom_belimlight = mainForm.wsIntercom.Cells(mainForm.adr_intercom_belimlight)
        mainForm.rng_intercom_PRLighting = mainForm.wsIntercom.Cells(mainForm.adr_intercom_PRLighting)
        mainForm.rng_intercom_blackout = mainForm.wsIntercom.Cells(mainForm.adr_intercom_blackout)
        mainForm.rng_intercom_vision = mainForm.wsIntercom.Cells(mainForm.adr_intercom_vision)
    End Sub


End Module
