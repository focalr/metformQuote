Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO

Public Class Form1

    Private Const setupLaborMarkup As Decimal = 1.33
    Private Const H13pricePerPound As Decimal = 1.8
    Private Const H43pricePerPound As Decimal = 11.75
    Private Const M2pricePerPound As Decimal = 5.5
    Private Const materialMarkup As Decimal = 1.2
    Private Const H13heatTreat As Decimal = 0.5
    Private Const H43heatTreat As Decimal = 1.0
    Private Const M2heatTreat As Decimal = 2.0
    Private Const H13nitride As Decimal = 2.0
    Private Const H43nitride As Decimal = 2.0
    Private Const M2nitride As Decimal = 2.0

    Private arrayH13 As Double() = {0.17606, 0.22899, 0.35567, 0.42944, 0.51353, 0.6015, 0.6963,
                                   0.93407, 1.033, 1.1656, 1.287, 1.4453, 1.74258,
                                   2.0676, 2.4204, 2.8011, 3.17586, 3.66377, 3.84, 4.0716,
                                   4.315, 4.62175, 5.1424, 5.60123, 6.2, 6.89584,
                                   7.52875, 8.24348, 9.65266, 11.01594, 12.03424, 13.04762}
    Private arrayH43 As Double() = {1.14814, 1.31707, 1.44444, 2.61, 2.85, 3.02, 3.19267, 3.93825, 3.94782, 4.406, 4.63188, 6.933}
    Private arrayM2 As Double() = {0.05895, 0.22899, 0.28886, 0.634,
                                         1.4453, 2.0676, 2.42049, 2.875, 3.17586, 3.22984, 4.0716,
                                         5.22, 6.87727, 7.56759}
    Private array As Array
    Private materialPrice As String
    Private heatTreatPrice As String
    Private nitridePrice As String
    Private index As Integer = 0
    Private purchaseOrder As String

    Private file_location As String = "I:\Metform Quotes\Metform Quotes.xlsx"
    Private printout_location As String = "I:\Metform Quotes\"

    Private xlApp As Excel.Application
    Private xlWorkbook As Excel.Workbook
    Private xlWorksheet As Excel.Worksheet

    Private quoteDay As Date = Date.Now

    Private table1 As DataTable


    ' form loading commands

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lbSteel.SelectedItem = "H13"

        txtQuantity.Select()

        ' bar stock tab objects
        ' loads form with rbno checked
        rbno.Checked = True

        ' loads form txtboxes with initial
        txtcutoff.Text = ".19"
        txtgripstock.Text = "2."
        txtxtrastock.Text = ".05"
        txtBarLength.Text = "151"
        txtOrderQty.Text = ""

    End Sub

    ' form closing commands

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        ' closes word app if form is closed
        'oWord.close()

    End Sub


    ' buttons

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' adds contents to lower calculations

        ' Material row
        txtMatSteel.Text = lbSteelOD.SelectedItem.ToString + " " + lbSteel.SelectedItem.ToString
        txtMatLength.Text = txtpartLength.Text

        Dim i = lbSteelOD.SelectedIndex
        txtMatNum.Text = array(i).ToString

        txtMatPounds.Text = Decimal.Parse(txtMatLength.Text) * Decimal.Parse(txtMatNum.Text)
        txtMatPounds.Text = Math.Round(Decimal.Parse(txtMatPounds.Text), 2, MidpointRounding.AwayFromZero)
        txtMatPrice.Text = materialPrice
        txtMatPricePer.Text = Decimal.Parse(txtMatPounds.Text) * Decimal.Parse(txtMatPrice.Text)
        txtMatPricePer.Text = Math.Round(Decimal.Parse(txtMatPricePer.Text), 2, MidpointRounding.AwayFromZero)
        txtMatMarkup.Text = materialMarkup.ToString
        txtMatTotal.Text = Decimal.Parse(txtMatPricePer.Text) * Decimal.Parse(txtMatMarkup.Text)
        txtMatTotal.Text = Math.Round(Decimal.Parse(txtMatTotal.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Set-up row
        ' checks the length of SU for alignment
        If txtSU.Text.Length < 3 Then
            txtSU.Text = "   " + txtSU.Text
        Else
            txtSU.Text = txtSU.Text
        End If
        txtSU.Text = txtSetup.Text
        txtSU_1.Text = setupLaborMarkup
        txtEquals.Text = Decimal.Parse(txtSU.Text) * Decimal.Parse(txtSU_1.Text)
        txtEquals.Text = Math.Round(Decimal.Parse(txtEquals.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")
        txtpartQty.Text = txtQuantity.Text
        txtTotal.Text = Decimal.Parse(txtEquals.Text) / Decimal.Parse(txtpartQty.Text)
        txtTotal.Text = Math.Round(Decimal.Parse(txtTotal.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Labor row
        txtL.Text = txtLabor.Text
        txtL_1.Text = setupLaborMarkup
        txtTotal_1.Text = Decimal.Parse(txtL.Text) * Decimal.Parse(txtL_1.Text)
        txtTotal_1.Text = Math.Round(Decimal.Parse(txtTotal_1.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Material row
        txtTotal_2.Text = txtMatTotal.Text
        txtTotal_2.Text = Math.Round(Decimal.Parse(txtTotal_2.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Heat Treat row
        txtHT.Text = txtMatPounds.Text
        txtHT_1.Text = heatTreatPrice
        txtTotal_3.Text = Decimal.Parse(txtHT.Text) * Decimal.Parse(txtHT_1.Text)
        txtTotal_3.Text = Math.Round(Decimal.Parse(txtTotal_3.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Nitride row
        txtN.Text = txtMatPounds.Text
        txtN_1.Text = nitridePrice
        txtTotal_4.Text = Decimal.Parse(txtN.Text) * Decimal.Parse(txtN_1.Text)
        txtTotal_4.Text = Math.Round(Decimal.Parse(txtTotal_4.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Totals

        If txtTotal_4.Enabled = True Then

            txtTotals.Text = Decimal.Parse(txtTotal.Text) + Decimal.Parse(txtTotal_1.Text) + Decimal.Parse(txtTotal_2.Text) _
                                   + Decimal.Parse(txtTotal_3.Text) + Decimal.Parse(txtTotal_4.Text)
            txtTotals.Text = Math.Round(Decimal.Parse(txtTotals.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        Else

            txtTotals.Text = Decimal.Parse(txtTotal.Text) + Decimal.Parse(txtTotal_1.Text) + Decimal.Parse(txtTotal_2.Text) _
                                 + Decimal.Parse(txtTotal_3.Text)
            txtTotals.Text = Math.Round(Decimal.Parse(txtTotals.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        End If




    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        ' clears content boxes

        cbNitride.Checked = False
        txtLabor.Clear()
        txtPartName.Clear()
        txtpartType.Clear()
        txtQuantity.Clear()
        txtSetup.Clear()
        lbSteelOD.ClearSelected()
        lbSteel.SelectedItem = "H13"
        txtpartLength.Clear()


    End Sub

    Private Sub btnSaveToList_Click(sender As Object, e As EventArgs) Handles btnSaveToList.Click

        ' adds the contents of the part info into the list view, moves to next row
        Dim str(8) As String
        Dim itm As ListViewItem
        str(0) = txtPartName.Text
        str(1) = txtpartType.Text
        str(2) = lbSteel.SelectedItem.ToString
        str(3) = lbSteelOD.SelectedItem.ToString
        str(4) = txtpartQty.Text
        str(5) = txtSU.Text
        str(6) = txtL.Text

        ' checks for nitride
        ' sends for text file
        If cbNitride.Checked = True Then
            str(7) = "Y"
            Call nitride_yes()
        Else
            str(7) = Nothing
            Call nitride_no()
        End If

        str(8) = txtTotals.Text


        itm = New ListViewItem(str)
        ListView1.Items.Insert(index, itm)
        index += 1

        btnClear.PerformClick()
        btnClearBottom.PerformClick()

        txtQuantity.Select()



    End Sub

    Private Sub btnClearBottom_Click(sender As Object, e As EventArgs) Handles btnClearBottom.Click
        ' clears the txt boxes

        Dim boxes As TextBox = Nothing
        For Each con As Object In Me.gbCalculations.Controls
            If TypeOf con Is TextBox Then
                boxes = con
                boxes.Text = String.Empty
            End If
        Next


    End Sub

    Private Sub btnClearLV_Click(sender As Object, e As EventArgs) Handles btnClearLV.Click
        ' clears listview
        ' resets index variable back to zero

        index = 0

        ListView1.Items.Clear()

        tbControl.SelectedTab = tbPartBuilder
        txtQuantity.Select()

    End Sub

    Private Sub btnSaveQuote_Click(sender As Object, e As EventArgs) Handles btnSaveQuote.Click
        ' opens excel workbook
        ' populates listview items into excel sheet
        ' saves and closes workbook


        Call open_excel()

        ' gets to the last row + 1 for space
        Dim rownum As Integer = xlWorksheet.UsedRange.Rows.Count + 2
        Dim colnum As Integer = 2

        For Each item As ListViewItem In ListView1.Items
            xlWorksheet.Cells(rownum, 1) = Date.Now.ToString("MM-dd-yyyy")
            For i As Integer = 0 To item.SubItems.Count - 1
                xlWorksheet.Cells(rownum, colnum) = item.SubItems(i).Text
                colnum = colnum + 1
            Next
            rownum += 1
            colnum = 2
        Next

        Call close_excel()

        Dim now As String = quoteDay.ToString("MM-dd-yyyy")
        Dim file As String = printout_location & now & ".doc"
        System.IO.File.WriteAllText(file, txtprintout.Text)

        btnClearLV.PerformClick()




    End Sub

    Private Sub btnRecalc_Click(sender As Object, e As EventArgs) Handles btnRecalc.Click


        txtMatPounds.Text = Decimal.Parse(txtMatLength.Text) * Decimal.Parse(txtMatNum.Text)
        txtMatPounds.Text = Math.Round(Decimal.Parse(txtMatPounds.Text), 2, MidpointRounding.AwayFromZero)
        txtMatPricePer.Text = Decimal.Parse(txtMatPounds.Text) * Decimal.Parse(txtMatPrice.Text)
        txtMatPricePer.Text = Math.Round(Decimal.Parse(txtMatPricePer.Text), 2, MidpointRounding.AwayFromZero)
        txtMatTotal.Text = Decimal.Parse(txtMatPricePer.Text) * Decimal.Parse(txtMatMarkup.Text)
        txtMatTotal.Text = Math.Round(Decimal.Parse(txtMatTotal.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Set-up row
        ' checks the length of SU for alignment
        If txtSU.Text.Length < 3 Then
            txtSU.Text = "   " + txtSU.Text
        Else
            txtSU.Text = txtSU.Text
        End If

        txtEquals.Text = Decimal.Parse(txtSU.Text) * Decimal.Parse(txtSU_1.Text)
        txtEquals.Text = Math.Round(Decimal.Parse(txtEquals.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")
        txtTotal.Text = Decimal.Parse(txtEquals.Text) / Decimal.Parse(txtpartQty.Text)
        txtTotal.Text = Math.Round(Decimal.Parse(txtTotal.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Labor row
        txtTotal_1.Text = Decimal.Parse(txtL.Text) * Decimal.Parse(txtL_1.Text)
        txtTotal_1.Text = Math.Round(Decimal.Parse(txtTotal_1.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Material row
        txtTotal_2.Text = txtMatTotal.Text
        txtTotal_2.Text = Math.Round(Decimal.Parse(txtTotal_2.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Heat Treat row
        txtTotal_3.Text = Decimal.Parse(txtHT.Text) * Decimal.Parse(txtHT_1.Text)
        txtTotal_3.Text = Math.Round(Decimal.Parse(txtTotal_3.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Nitride row
        txtTotal_4.Text = Decimal.Parse(txtN.Text) * Decimal.Parse(txtN_1.Text)
        txtTotal_4.Text = Math.Round(Decimal.Parse(txtTotal_4.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        ' Totals

        If txtTotal_4.Enabled = True Then

            txtTotals.Text = Decimal.Parse(txtTotal.Text) + Decimal.Parse(txtTotal_1.Text) + Decimal.Parse(txtTotal_2.Text) _
                                   + Decimal.Parse(txtTotal_3.Text) + Decimal.Parse(txtTotal_4.Text)
            txtTotals.Text = Math.Round(Decimal.Parse(txtTotals.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        Else

            txtTotals.Text = Decimal.Parse(txtTotal.Text) + Decimal.Parse(txtTotal_1.Text) + Decimal.Parse(txtTotal_2.Text) _
                                 + Decimal.Parse(txtTotal_3.Text)
            txtTotals.Text = Math.Round(Decimal.Parse(txtTotals.Text), 1, MidpointRounding.AwayFromZero).ToString("F2")

        End If

    End Sub



    ' list box commands

    Private Sub lbSteel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbSteel.SelectedIndexChanged

        ' populate steel od by the steel type selected

        Select Case True

            Case lbSteel.SelectedItem = "H13"
                lbSteelOD.Items.Clear()
                array = arrayH13
                materialPrice = H13pricePerPound.ToString
                heatTreatPrice = H13heatTreat.ToString
                nitridePrice = H13nitride.ToString
                lbSteelOD.Items.Add(".87500")
                lbSteelOD.Items.Add("1.0000")
                lbSteelOD.Items.Add("1.2500")
                lbSteelOD.Items.Add("1.3750")
                lbSteelOD.Items.Add("1.5000")
                lbSteelOD.Items.Add("1.6250")
                lbSteelOD.Items.Add("1.7500")
                lbSteelOD.Items.Add("2.0000")
                lbSteelOD.Items.Add("2.1250")
                lbSteelOD.Items.Add("2.2500")
                lbSteelOD.Items.Add("2.3750")
                lbSteelOD.Items.Add("2.5000")
                lbSteelOD.Items.Add("2.7500")
                lbSteelOD.Items.Add("3.0000")
                lbSteelOD.Items.Add("3.2500")
                lbSteelOD.Items.Add("3.5000")
                lbSteelOD.Items.Add("3.7500")
                lbSteelOD.Items.Add("4.0000")
                lbSteelOD.Items.Add("4.1250")
                lbSteelOD.Items.Add("4.2500")
                lbSteelOD.Items.Add("4.3750")
                lbSteelOD.Items.Add("4.5000")
                lbSteelOD.Items.Add("4.7500")
                lbSteelOD.Items.Add("5.0000")
                lbSteelOD.Items.Add("5.2500")
                lbSteelOD.Items.Add("5.5000")
                lbSteelOD.Items.Add("5.7500")
                lbSteelOD.Items.Add("6.0000")
                lbSteelOD.Items.Add("6.5000")
                lbSteelOD.Items.Add("7.0000")
                lbSteelOD.Items.Add("7.2500")
                lbSteelOD.Items.Add("7.5000")

            Case lbSteel.SelectedItem = "H43"
                lbSteelOD.Items.Clear()
                array = arrayH43
                materialPrice = H43pricePerPound.ToString
                heatTreatPrice = H43heatTreat.ToString
                nitridePrice = H43nitride.ToString
                lbSteelOD.Items.Add("2.2700")
                lbSteelOD.Items.Add("2.4100")
                lbSteelOD.Items.Add("2.5200")
                lbSteelOD.Items.Add("3.4200")
                lbSteelOD.Items.Add("3.5400")
                lbSteelOD.Items.Add("3.6700")
                lbSteelOD.Items.Add("3.7900")
                lbSteelOD.Items.Add("4.2020")
                lbSteelOD.Items.Add("4.2120")
                lbSteelOD.Items.Add("4.4400")
                lbSteelOD.Items.Add("4.5650")
                lbSteelOD.Items.Add("5.5850")


            Case lbSteel.SelectedItem = "M2"
                lbSteelOD.Items.Clear()
                array = arrayM2
                materialPrice = M2pricePerPound.ToString
                heatTreatPrice = M2heatTreat.ToString
                nitridePrice = M2nitride.ToString
                lbSteelOD.Items.Add(".50000")
                lbSteelOD.Items.Add("1.0000")
                lbSteelOD.Items.Add("1.1250")
                lbSteelOD.Items.Add("1.6875")
                lbSteelOD.Items.Add("2.5000")
                lbSteelOD.Items.Add("3.0000")
                lbSteelOD.Items.Add("3.2500")
                lbSteelOD.Items.Add("3.5625")
                lbSteelOD.Items.Add("3.7500")
                lbSteelOD.Items.Add("3.8120")
                lbSteelOD.Items.Add("4.2500")
                lbSteelOD.Items.Add("4.7500")
                lbSteelOD.Items.Add("5.5625")
                lbSteelOD.Items.Add("5.8125")


        End Select


    End Sub

    ' checkbox commands

    Private Sub cbNitride_CheckedChanged(sender As Object, e As EventArgs) Handles cbNitride.CheckedChanged
        ' enables or disables Nitride boxes and labels as clicked

        If cbNitride.Checked = True Then
            lblN.Enabled = True
            txtN.Enabled = True
            lblX_4.Enabled = True
            txtN_1.Enabled = True
            txtTotal_4.Enabled = True
        Else
            lblN.Enabled = False
            txtN.Enabled = False
            lblX_4.Enabled = False
            txtN_1.Enabled = False
            txtTotal_4.Enabled = False
        End If


    End Sub


    ' custom functions

    Private Function nitride_yes()

        txtprintout.Text = txtprintout.Text &
          "(" + txtQuantity.Text & ")" & "   " + txtPartName.Text & "   " + lbSteel.SelectedItem.ToString & vbNewLine &
         vbNewLine &
        vbTab & "SU = " + txtSU.Text & " X " + txtSU_1.Text & " = " + txtEquals.Text & " / " + txtpartQty.Text & vbTab & vbTab &
                                                   "=" & vbTab & "$" + txtTotal.Text &
      vbNewLine &
      vbTab & "L = " + txtL.Text & " X " + txtL_1.Text & vbTab & vbTab & vbTab & vbTab & vbTab & "      " + txtTotal_1.Text &
         vbNewLine &
        vbTab & "M = " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "      " + txtTotal_2.Text &
       vbNewLine &
      vbTab & "HT = " & txtHT.Text & " X " + txtHT_1.Text & vbTab & vbTab & vbTab & vbTab & vbTab & "      " + txtTotal_3.Text &
      vbNewLine &
      vbTab & "N = " + txtN.Text & " X " + txtN_1.Text & vbTab & vbTab & vbTab & vbTab & vbTab & "      " + txtTotal_4.Text &
      vbNewLine &
      vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "---------------" &
      vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "$" + txtTotals.Text &
       vbNewLine & vbNewLine &
       "MAT= " + lbSteel.SelectedItem.ToString & " " + lbSteelOD.SelectedItem.ToString & " @ " + txtpartLength.Text &
           " X " + txtMatNum.Text & " = " + txtMatPounds.Text & " @ " + txtMatPrice.Text & " = " + txtMatPricePer.Text &
           " X " + txtMatMarkup.Text & " = " + txtMatTotal.Text &
           vbNewLine &
      "-----------------------------------------------------------------" &
       vbNewLine &
       vbNewLine

        Return txtprintout.Text

    End Function

    Private Function nitride_no()
        txtprintout.Text = txtprintout.Text &
          "(" + txtQuantity.Text & ")" & "   " + txtPartName.Text & "   " + lbSteel.SelectedItem.ToString & vbNewLine &
         vbNewLine &
        vbTab & "SU = " + txtSU.Text & " X " + txtSU_1.Text & " = " + txtEquals.Text & " / " + txtpartQty.Text & vbTab & vbTab &
                                                   "=" & vbTab & "$" + txtTotal.Text &
      vbNewLine &
      vbTab & "L = " + txtL.Text & " X " + txtL_1.Text & vbTab & vbTab & vbTab & vbTab & vbTab & "      " + txtTotal_1.Text &
         vbNewLine &
        vbTab & "M = " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "      " + txtTotal_2.Text &
       vbNewLine &
      vbTab & "HT = " & txtHT.Text & " X " + txtHT_1.Text & vbTab & vbTab & vbTab & vbTab & vbTab & "      " + txtTotal_3.Text &
      vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "---------------" &
      vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "$" + txtTotals.Text &
       vbNewLine & vbNewLine &
       "MAT= " + lbSteel.SelectedItem.ToString & " " + lbSteelOD.SelectedItem.ToString & " @ " + txtpartLength.Text &
           " X " + txtMatNum.Text & " = " + txtMatPounds.Text & " @ " + txtMatPrice.Text & " = " + txtMatPricePer.Text &
           " X " + txtMatMarkup.Text & " = " + txtMatTotal.Text &
           vbNewLine &
      "-----------------------------------------------------------------" &
       vbNewLine &
       vbNewLine

        Return txtPrintout.Text

    End Function

    

    


    

    ' excel operations

    Private Sub open_excel()

        ' opens excel employee workbook
        xlApp = New Excel.Application
        xlWorkbook = xlApp.Workbooks.Open(file_location, Notify:=False, [ReadOnly]:=False)
        xlWorksheet = xlWorkbook.Worksheets("Sheet1")

    End Sub

    Private Sub close_excel()

        xlWorkbook.Save()
        xlWorkbook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkbook)
        releaseObject(xlWorksheet)

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub



    ' bar stock tab code

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click

        Call Calculations()

    End Sub

    Private Sub Calculations()

        ' calculates proper lengths

        'set variables for inputs
        Dim pl As Decimal
        Decimal.TryParse(txtpart_length_barstock.Text, pl)

        Dim co As Decimal
        Decimal.TryParse(txtcutoff.Text, co)

        Dim xs As Decimal
        Decimal.TryParse(txtxtrastock.Text, xs)

        Dim gs As Decimal
        Decimal.TryParse(txtgripstock.Text, gs)

        Dim bl As Decimal
        Decimal.TryParse(txtBarLength.Text, bl)

        'adjust waycenter by radiobutton
        Dim wc As Decimal
        If rbyes.Checked = True Then
            wc = 0.25
        Else
            wc = 0.0
        End If

        'part length calc
        Dim part_total As Decimal = pl + co + xs + wc

        'parts per bar calc
        Dim parts_per_cut As Integer
        Dim length_per_cut As Decimal = 0.0

        ' finds # of parts per <45" bar length
        ' gives _parts as # of parts per cut
        ' gives length_per_cut bar length for that many parts
        Dim _parts = 1
        Do Until length_per_cut > 45
            length_per_cut = (part_total * _parts) + gs
            _parts += 1
        Loop
        ' adjustment
        _parts -= 1
        length_per_cut = (part_total * _parts) + gs
        parts_per_cut = _parts

        ' begin subtracting # of cuts per full bar length
        Dim cuts_per_fullbar As Integer
        Dim bar_length As Decimal = bl
        Dim _cuts = 1

        ' subtracts length of cut from full bar length until less than single length per cut
        Do Until bar_length < length_per_cut
            bar_length = bar_length - length_per_cut
            _cuts += 1
        Loop
        ' adjustment
        cuts_per_fullbar = _cuts - 1

        ' find # of parts left with the remaining piece of bar left

        Dim remaining_parts As Integer = 1
        Dim remaining_barlenght As Decimal = bar_length
        Do Until remaining_barlenght < (part_total + gs)
            remaining_barlenght = remaining_barlenght - (part_total + gs)
            remaining_parts += 1
        Loop
        ' adjustment
        remaining_parts -= 1


        ' math to determine total # of parts per full bar
        Dim total_parts_per_fullbar As Integer

        total_parts_per_fullbar = (parts_per_cut * cuts_per_fullbar) + remaining_parts

        txtPartsPerBar.Text = total_parts_per_fullbar.ToString

        If txtOrderQty.Text IsNot "" Then
            ' calc the total inches needed for order
            Dim orderQty As Integer
            Integer.TryParse(txtOrderQty.Text, orderQty)
            Dim a As Decimal = orderQty / total_parts_per_fullbar
            Dim totalInches As Decimal
            totalInches = a * bl

            txtTotalInches.Text = totalInches.ToString
            txtTotalInches.Text = Math.Round(Decimal.Parse(txtTotalInches.Text), 2, MidpointRounding.AwayFromZero)
        End If


        'creating a new datatable with column names
        table1 = New DataTable
        table1.Columns.Add("# Parts")
        table1.Columns.Add("Length")

        'calculating total lengths
        'populating length column rows
        Dim total_length As Decimal = 0.0
        Dim i As Integer = 1
        While total_length <= 60
            total_length = part_total * i + gs
            table1.Rows.Add(i, total_length)

            txtGuts.Text =
                txtGuts.Text + i.ToString & vbTab + total_length.ToString & vbNewLine

            i += 1
        End While

        txtoutput.Text = txtoutput.Text & vbNewLine + txtpart_length_barstock.Text & vbNewLine + txtGuts.Text

        'assign table data to datagridview
        dgvresults.DataSource = table1

        ' set column width for all columns
        For Each c As DataGridViewColumn In dgvresults.Columns
            c.Width = 120
        Next

    End Sub

    Private Sub btnClear_BarStock_Click(sender As Object, e As EventArgs) Handles btnClear_BarStock.Click

        ' clears part length box
        ' resets radiobuttons to no
        ' resets txtboxes to inital on load
        rbno.Checked = True

        txtcutoff.Text = ".125"
        txtgripstock.Text = "2."
        txtxtrastock.Text = ".05"
        txtpart_length_barstock.Text = Nothing
        txtPartsPerBar.Clear()
        txtOrderQty.Text = ""
        txtTotalInches.Clear()

        ' clears datagridview
        dgvresults.DataSource = Nothing

        ' puts cursor in part length txtbox
        txtpart_length_barstock.Focus()

        ' clears txtouputs
        txtoutput.Clear()

    End Sub

    Private Sub txtpart_length_barstock_KeyUp(sender As Object, e As KeyEventArgs) Handles txtpart_length_barstock.KeyUp,
       txtxtrastock.KeyUp, txtcutoff.KeyUp, txtgripstock.KeyUp, txtOrderQty.KeyUp, rbno.KeyUp, rbyes.KeyUp

        If (e.KeyCode = Keys.Escape) Then
            btnClear_BarStock.PerformClick()
        End If

        If (e.KeyCode = Keys.Enter) Then
            btnCalc.PerformClick()
        End If

    End Sub

    Private Sub btnCalculator_Click(sender As Object, e As EventArgs) Handles btnCalculator.Click
        ' clicking this button takes you to bar stock tab

        tbControl.SelectedTab = tbpBarStock



    End Sub
End Class
