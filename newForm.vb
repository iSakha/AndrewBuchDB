Imports OfficeOpenXml
Imports OfficeOpenXml.Table
Imports System.IO

Public Class newForm
    Public txtName() As String
    Public txtQty() As String
    Public qty_belimlight1, qty_belimlight2, qty_belimlight3 As Integer
    Public qty_belimlight As Integer
    Public qty_PRlighting1, qty_PRlighting2, qty_PRlighting3 As Integer
    Public qty_PRlighting As Integer
    Public qty_blackout1, qty_blackout2, qty_blackout3 As Integer
    Public qty_blackout As Integer
    Public qty_vision1, qty_vision2, qty_vision3 As Integer
    Public qty_vision As Integer
    Public qty_stage1, qty_stage2, qty_stage3 As Integer
    Public qty_stage As Integer
    Private Sub addNewItemForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'txtName = {
        '    txt_belimlight1_addform.Text, txt_belimlight2_addform.Text, txt_belimlight3_addform.Text,
        '    txt_PRlighting1_addform.Text, txt_PRlighting2_addform.Text, txt_PRlighting3_addform.Text,
        '    txt_blackout1_addform.Text, txt_blackout2_addform.Text, txt_blackout3_addform.Text,
        '    txt_vision1_addform.Text, txt_vision2_addform.Text, txt_vision3_addform.Text,
        '    txt_stage1_addform.Text, txt_stage2_addform.Text, txt_stage3_addform.Text
        '    }

        'txtQty = {txt_qty_belimlight1_addform.Text, txt_qty_belimlight2_addform.Text, txt_qty_belimlight3_addform.Text,
        '    txt_qty_PRlighting1_addform.Text, txt_qty_PRlighting2_addform.Text, txt_qty_PRlighting3_addform.Text,
        '    txt_qty_blackout1_addform.Text, txt_qty_blackout2_addform.Text, txt_qty_blackout3_addform.Text,
        '    txt_qty_vision1_addform.Text, txt_qty_vision2_addform.Text, txt_qty_vision3_addform.Text,
        '    txt_qty_stage1_addform.Text, txt_qty_stage2_addform.Text, txt_qty_stage3_addform.Text
        '}
    End Sub
    '===================================================================================
    '             === ADD data to DB ===
    '===================================================================================
    Private Sub btn_add_addform_Click(sender As Object, e As EventArgs) Handles btn_add_addform.Click

        Dim i As Integer

        i = mainForm.selectedCategoryIndex
        qty_belimlight1 = Integer.Parse(txt_qty_belimlight1_addform.Text)
        qty_belimlight2 = Integer.Parse(txt_qty_belimlight2_addform.Text)
        qty_belimlight3 = Integer.Parse(txt_qty_belimlight3_addform.Text)

        qty_belimlight = qty_belimlight1 + qty_belimlight2 + qty_belimlight3

        qty_PRlighting1 = Integer.Parse(txt_qty_PRlighting1_addform.Text)
        qty_PRlighting2 = Integer.Parse(txt_qty_PRlighting2_addform.Text)
        qty_PRlighting3 = Integer.Parse(txt_qty_PRlighting3_addform.Text)

        qty_PRlighting = qty_PRlighting1 + qty_PRlighting2 + qty_PRlighting3

        qty_blackout1 = Integer.Parse(txt_qty_blackout1_addform.Text)
        qty_blackout2 = Integer.Parse(txt_qty_blackout2_addform.Text)
        qty_blackout3 = Integer.Parse(txt_qty_blackout3_addform.Text)

        qty_blackout = qty_blackout1 + qty_blackout2 + qty_blackout3

        qty_vision1 = Integer.Parse(txt_qty_vision1_addform.Text)
        qty_vision2 = Integer.Parse(txt_qty_vision2_addform.Text)
        qty_vision3 = Integer.Parse(txt_qty_vision3_addform.Text)

        qty_vision = qty_vision1 + qty_vision2 + qty_vision3

        qty_stage1 = Integer.Parse(txt_qty_stage1_addform.Text)
        qty_stage2 = Integer.Parse(txt_qty_stage2_addform.Text)
        qty_stage3 = Integer.Parse(txt_qty_stage3_addform.Text)

        qty_stage = qty_stage1 + qty_stage2 + qty_stage3

        txt_qty_belimlight.Text = qty_belimlight
        txt_qty_PRlighting.Text = qty_PRlighting
        txt_qty_blackout.Text = qty_blackout
        txt_qty_vision.Text = qty_vision
        txt_qty_stage.Text = qty_stage

        addData(i)
    End Sub

    Private Sub cmb_category_addform_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_category_addform.SelectedIndexChanged
        mainForm.selectedCategoryIndex = cmb_category_addform.SelectedIndex
    End Sub
    Private Sub btn_close_addform_Click(sender As Object, e As EventArgs) Handles btn_close_addform.Click
        Me.Close()
    End Sub
    ''           Quantity Belimlight
    'Private Sub txt_qty_belimlight1_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_belimlight1_addform.TextChanged
    '    'Dim qty_belimlight As Integer
    '    'qty_belimlight = CInt(txt_qty_belimlight1_addform.Text) + CInt(txt_qty_belimlight2_addform.Text) +
    '    '    CInt(txt_qty_belimlight3_addform.Text)
    '    'txt_qty_belimlight.Text = CStr(qty_belimlight)
    'End Sub

    'Private Sub txt_qty_belimlight2_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_belimlight2_addform.TextChanged

    '    Dim qty_belimlight1, qty_belimlight2, qty_belimlight3 As Integer
    '    Dim qty_belimlight As Integer

    '    qty_belimlight1 = Integer.Parse(txt_qty_belimlight1_addform.Text)
    '    qty_belimlight2 = Integer.Parse(txt_qty_belimlight2_addform.Text)
    '    qty_belimlight3 = Integer.Parse(txt_qty_belimlight3_addform.Text)

    '    'qty_belimlight = Convert.ToInt32(qty_belimlight1) + Convert.ToInt32(qty_belimlight2) + Convert.ToInt32(qty_belimlight3)

    '    'qty_belimlight = CInt(txt_qty_belimlight1_addform.Text) + CInt(txt_qty_belimlight2_addform.Text) +
    '    '    CInt(txt_qty_belimlight3_addform.Text)
    '    'txt_qty_belimlight.Text = CStr(qty_belimlight)
    'End Sub

    ''Private Sub txt_qty_belimlight3_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_belimlight3_addform.TextChanged
    ''    Dim qty_belimlight As Integer
    ''    qty_belimlight = CInt(txt_qty_belimlight1_addform.Text) + CInt(txt_qty_belimlight2_addform.Text) +
    ''        CInt(txt_qty_belimlight3_addform.Text)
    ''    txt_qty_belimlight.Text = qty_belimlight
    ''End Sub

    ''Private Sub txt_qty_PRlighting1_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_PRlighting1_addform.TextChanged
    ''    'txt_qty_PRlighting.Text = CInt(txt_qty_PRlighting1_addform.Text) + CInt(txt_qty_PRlighting2_addform.Text) +
    ''    '    CInt(txt_qty_PRlighting3_addform.Text)
    ''End Sub

    'Private Sub txt_qty_PRlighting2_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_PRlighting2_addform.TextChanged
    '    txt_qty_PRlighting.Text = Integer.Parse(txt_qty_PRlighting1_addform.Text) + Integer.Parse(txt_qty_PRlighting2_addform.Text) +
    '        Integer.Parse(txt_qty_PRlighting3_addform.Text)
    'End Sub

    'Private Sub txt_qty_PRlighting3_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_PRlighting3_addform.TextChanged
    '    txt_qty_PRlighting.Text = CInt(txt_qty_PRlighting1_addform.Text) + CInt(txt_qty_PRlighting2_addform.Text) +
    '        CInt(txt_qty_PRlighting3_addform.Text)
    'End Sub
    ''           Quantity Blackout
    'Private Sub txt_qty_blackout1_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_blackout1_addform.TextChanged
    '    'txt_qty_blackout.Text = CInt(txt_qty_blackout1_addform.Text) + CInt(txt_qty_blackout2_addform.Text) +
    '    '    CInt(txt_qty_blackout3_addform.Text)
    'End Sub

    'Private Sub txt_qty_blackout2_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_blackout2_addform.TextChanged
    '    txt_qty_blackout.Text = CInt(txt_qty_blackout1_addform.Text) + CInt(txt_qty_blackout2_addform.Text) +
    '        CInt(txt_qty_blackout3_addform.Text)
    'End Sub

    'Private Sub txt_qty_blackout3_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_blackout3_addform.TextChanged
    '    txt_qty_blackout.Text = CInt(txt_qty_blackout1_addform.Text) + CInt(txt_qty_blackout2_addform.Text) +
    '        CInt(txt_qty_blackout3_addform.Text)
    'End Sub
    ''           Quantity Vision
    'Private Sub txt_qty_vision1_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_vision1_addform.TextChanged
    '    'txt_qty_vision.Text = CInt(txt_qty_vision1_addform.Text) + CInt(txt_qty_vision2_addform.Text) +
    '    '    CInt(txt_qty_vision3_addform.Text)
    'End Sub

    'Private Sub txt_qty_vision2_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_vision2_addform.TextChanged
    '    txt_qty_vision.Text = CInt(txt_qty_vision1_addform.Text) + CInt(txt_qty_vision2_addform.Text) +
    '        CInt(txt_qty_vision3_addform.Text)
    'End Sub

    'Private Sub txt_qty_vision3_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_vision3_addform.TextChanged
    '    txt_qty_vision.Text = CInt(txt_qty_vision1_addform.Text) + CInt(txt_qty_vision2_addform.Text) +
    '        CInt(txt_qty_vision3_addform.Text)
    'End Sub
    ''           Quantity Stage
    'Private Sub txt_qty_stage1_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_stage1_addform.TextChanged
    '    'txt_qty_stage.Text = CInt(txt_qty_stage1_addform.Text) + CInt(txt_qty_stage2_addform.Text) +
    '    '    CInt(txt_qty_stage3_addform.Text)
    'End Sub

    'Private Sub txt_qty_stage2_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_stage2_addform.TextChanged
    '    txt_qty_stage.Text = CInt(txt_qty_stage1_addform.Text) + CInt(txt_qty_stage2_addform.Text) +
    '        CInt(txt_qty_stage3_addform.Text)
    'End Sub

    'Private Sub txt_qty_stage3_addform_TextChanged(sender As Object, e As EventArgs) Handles txt_qty_stage3_addform.TextChanged
    '    txt_qty_stage.Text = CInt(txt_qty_stage1_addform.Text) + CInt(txt_qty_stage2_addform.Text) +
    '        CInt(txt_qty_stage3_addform.Text)
    'End Sub



End Class