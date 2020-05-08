Module myFunctions

    Sub initMovHeads()
        mainForm.tbl_movHeads_belimlight = mainForm.tbl_Light_Collection.Item("movHeads_belimlight")
        mainForm.tbl_movHeads_PRLighting = mainForm.tbl_Light_Collection.Item("movHeads_PRlighting")
        mainForm.tbl_movHeads_blackout = mainForm.tbl_Light_Collection.Item("movHeads_blackout")
        mainForm.tbl_movHeads_vision = mainForm.tbl_Light_Collection.Item("movHeads_vision")

        mainForm.r_movHeads_belimlight = mainForm.tbl_movHeads_belimlight.Address.Rows
        mainForm.r_movHeads_PRLighting = mainForm.tbl_movHeads_PRLighting.Address.Rows
        mainForm.r_movHeads_blackout = mainForm.tbl_movHeads_blackout.Address.Rows
        mainForm.r_movHeads_vision = mainForm.tbl_movHeads_vision.Address.Rows

        mainForm.c_movHeads_belimlight = mainForm.tbl_movHeads_belimlight.Address.Columns
        mainForm.c_movHeads_PRLighting = mainForm.tbl_movHeads_PRLighting.Address.Columns
        mainForm.c_movHeads_blackout = mainForm.tbl_movHeads_blackout.Address.Columns
        mainForm.c_movHeads_vision = mainForm.tbl_movHeads_vision.Address.Columns

        mainForm.adr_movHeads_belimlight = mainForm.tbl_movHeads_belimlight.Address.Address
        mainForm.adr_movHeads_PRlighting = mainForm.tbl_movHeads_PRLighting.Address.Address
        mainForm.adr_movHeads_blackout = mainForm.tbl_movHeads_blackout.Address.Address
        mainForm.adr_movHeads_vision = mainForm.tbl_movHeads_vision.Address.Address

        mainForm.rng_movHeads_belimlight = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_belimlight)
        mainForm.rng_movHeads_PRlighting = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_PRlighting)
        mainForm.rng_movHeads_blackout = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_blackout)
        mainForm.rng_movHeads_vision = mainForm.wsMovHeads.Cells(mainForm.adr_movHeads_vision)
    End Sub

    Sub initStrobes()
        mainForm.tbl_strobes_belimlight = mainForm.tbl_Light_Collection.Item("strobes_belimlight")
        mainForm.tbl_strobes_PRLighting = mainForm.tbl_Light_Collection.Item("strobes_PRlighting")
        mainForm.tbl_strobes_blackout = mainForm.tbl_Light_Collection.Item("strobes_blackout")
        mainForm.tbl_strobes_vision = mainForm.tbl_Light_Collection.Item("strobes_vision")

        mainForm.r_strobes_belimlight = mainForm.tbl_strobes_belimlight.Address.Rows
        mainForm.r_strobes_PRLighting = mainForm.tbl_strobes_PRLighting.Address.Rows
        mainForm.r_strobes_blackout = mainForm.tbl_strobes_blackout.Address.Rows
        mainForm.r_strobes_vision = mainForm.tbl_strobes_vision.Address.Rows

        mainForm.c_strobes_belimlight = mainForm.tbl_strobes_belimlight.Address.Columns
        mainForm.c_strobes_PRLighting = mainForm.tbl_strobes_PRLighting.Address.Columns
        mainForm.c_strobes_blackout = mainForm.tbl_strobes_blackout.Address.Columns
        mainForm.c_strobes_vision = mainForm.tbl_strobes_vision.Address.Columns

        mainForm.adr_strobes_belimlight = mainForm.tbl_strobes_belimlight.Address.Address
        mainForm.adr_strobes_PRlighting = mainForm.tbl_strobes_PRLighting.Address.Address
        mainForm.adr_strobes_blackout = mainForm.tbl_strobes_blackout.Address.Address
        mainForm.adr_strobes_vision = mainForm.tbl_strobes_vision.Address.Address

        mainForm.rng_strobes_belimlight = mainForm.wsStrobes.Cells(mainForm.adr_strobes_belimlight)
        mainForm.rng_strobes_PRlighting = mainForm.wsStrobes.Cells(mainForm.adr_strobes_PRlighting)
        mainForm.rng_strobes_blackout = mainForm.wsStrobes.Cells(mainForm.adr_strobes_blackout)
        mainForm.rng_strobes_vision = mainForm.wsStrobes.Cells(mainForm.adr_strobes_vision)
    End Sub

End Module
