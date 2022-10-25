Imports System.Drawing.Text
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar

Public Class form1
    Dim Forsta As Double
    Dim Andra As Decimal
    Dim Operatorer As Integer
    Dim Operator_valjare As Boolean = False

    Private Sub form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Nollställer inmatningsrutan vid uppstart av appen
        txtInmatning.Text = ""

    End Sub

    Private Sub txtInmatning_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtInmatning.KeyPress
        'Tillåter punkt
        If e.KeyChar = "."c AndAlso InStr(sender.text, ".") > 0 Then
            e.Handled = True
        End If

        'Tillåter siffor
        If (e.KeyChar < "0"c OrElse e.KeyChar > "9"c) AndAlso e.KeyChar <> "."c AndAlso e.KeyChar <> vbBack Then
            e.Handled = True
        End If

    End Sub

    Private Sub txtInmatning_Click(sender As Object, e As EventArgs) Handles txtInmatning.Click, btn0.Click, btn1.Click, btn2.Click, btn3.Click, btn4.Click, btn5.Click, btn6.Click, btn7.Click, btn8.Click, btn9.Click
        'Begränsar att man bara kan skriva in 8 tal
        If txtInmatning.TextLength >= Int(8) Then
            MsgBox("Max 8 tecken")
        End If

    End Sub


    Private Sub btn0_Click(sender As Object, e As EventArgs) Handles btn0.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 0 till
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "0"

        End If

    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 1 till annars är det 1

        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "1"
        Else
            txtInmatning.Text = "1"
        End If
    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 2 till annars är det 2
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "2"
        Else
            txtInmatning.Text = "2"
        End If
    End Sub

    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 3 till annars är det 3
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "3"
        Else
            txtInmatning.Text = "3"
        End If
    End Sub

    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 4 till annars är det 4
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "4"
        Else
            txtInmatning.Text = "4"
        End If
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 5 till annars är det 5
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "5"
        Else
            txtInmatning.Text = "5"
        End If
    End Sub

    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 6 till annars är det 6
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "6"
        Else
            txtInmatning.Text = "6"
        End If

    End Sub

    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 7 till annars är det 7
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "7"
        Else
            txtInmatning.Text = "7"
        End If
    End Sub

    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 8 till annars är det 8
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "8"
        Else
            txtInmatning.Text = "8"
        End If
    End Sub

    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        'Om inmatningsrutan inte är lika med 0 så läggs 9 till annars är det 9
        If txtInmatning.Text <> "0" Then
            txtInmatning.Text += "9"
        Else
            txtInmatning.Text = "9"
        End If
    End Sub

    Private Sub btnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        'Om inmatningsrutan är tom så när man trycker på decimaltecknet läggs 0 till före
        If txtInmatning.Text = "" Then
            txtInmatning.Text += "0."
            'Annars läggs bara decimaltecknet till
        ElseIf Not (txtInmatning.Text.Contains(".")) Then
            txtInmatning.Text += "."

        End If

    End Sub

    Private Sub btnSumma_Click(sender As Object, e As EventArgs) Handles btnSumma.Click
        'Om man valt ett räknesätt så kan uträkningen ske
        If Operator_valjare = True Then
            Andra = txtInmatning.Text

            'Operatorn för addition
            If Operatorer = 1 Then
                'Avrundar svaret så det inte är mer än 8 tecken 
                txtInmatning.Text = Math.Round(Forsta + Andra, 7)

                'Operatorn för subtraktion
            ElseIf Operatorer = 2 Then
                'Avrundar svaret så det inte är mer än 8 tecken 
                txtInmatning.Text = Math.Round(Forsta - Andra, 7)

                'Operatorn för multiplikation
            ElseIf Operatorer = 3 Then
                'Avrundar svaret så det inte är mer än 8 tecken 
                txtInmatning.Text = Math.Round(Forsta * Andra, 7)

            Else
                'Om man försöker dela något med 0 så skrivs ett felmeddelande ut
                If Andra = 0 Then
                    txtInmatning.Text = "Error, kan inte dela med 0"
                Else
                    'Operatorn för divition
                    txtInmatning.Text = Math.Round(Forsta / Andra, 7)     'Avrundar svaret så det inte är mer än 8 tecken 
                End If
            End If
            Operator_valjare = False
        End If


    End Sub


    Private Sub btnRensaAllt_Click(sender As Object, e As EventArgs) Handles btnRensaAllt.Click

        'Rensar allt
        txtInmatning.Clear()


    End Sub

    Private Sub btnRensaInmatning_Click(sender As Object, e As EventArgs) Handles btnRensaInmatning.Click
        'Rensar inmatningsrutan
        txtInmatning.Text = ""

    End Sub

    Private Sub btnAddition_Click(sender As Object, e As EventArgs) Handles btnAddition.Click
        'Rätt operator väljs och texten blir till 0
        Forsta = Convert.ToDouble(txtInmatning.Text)
        txtInmatning.Text = "0"
        Operator_valjare = True
        Operatorer = 1



    End Sub

    Private Sub btnSubtraktion_Click(sender As Object, e As EventArgs) Handles btnSubtraktion.Click
        'Rätt operator väljs och texten blir till 0
        Forsta = txtInmatning.Text
        txtInmatning.Text = "0"
        Operator_valjare = True
        Operatorer = 2

    End Sub

    Private Sub btnMultiplikation_Click(sender As Object, e As EventArgs) Handles btnMultiplikation.Click
        'Rätt operator väljs och texten blir till 0
        Forsta = txtInmatning.Text
        txtInmatning.Text = "0"
        Operator_valjare = True
        Operatorer = 3
    End Sub

    Private Sub btnDivision_Click(sender As Object, e As EventArgs) Handles btnDivision.Click
        'Rätt operator väljs och texten blir till 0
        Forsta = txtInmatning.Text
        txtInmatning.Text = "0"
        Operator_valjare = True
        Operatorer = 4

    End Sub

End Class
