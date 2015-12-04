    'Sebastian Kaye
    'Mr. Scott
    'Basic Programming Period 6
    '1 December 2015

    'Calculates the monthly payment of a loan.

    'initializes the program and sets all the numbers and radio buttons
    Private Sub init()

        lnAmount.Maximum = 25000
        lnAmount.Value = 0
        intRate.Value = 5
        t1Years.Checked = False
        t2Years.Checked = False
        t3Years.Checked = False
        crdRate.Text = ""
        outMonthPay.Text = ""

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        init()

    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click

        init()

    End Sub

    Private Sub btnCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click

        '==========Post-Initialization==========

        'the variables in order: Interest rate, Adjusted interest rate, Payment plan, Payment per month, Error flag
        Dim dIntRate, dAdjRate, dPlanTier, dPayment As Double
        Dim errorMsg As Boolean = False

        'calculates the interest rate
        dIntRate = (Convert.ToDouble(intRate.Value) / 100) / 12

        'Adjusts the rate depending on your credit score
        Select Case crdRate.SelectedIndex
            Case 0 'credit rating of A
                dAdjRate = dIntRate
            Case 1 'credit rating of B
                dAdjRate = dIntRate * 1.1
            Case 2 'credit rating of C
                dAdjRate = dIntRate * 1.15
            Case 3 'credit rating of D
                dAdjRate = dIntRate * 1.17
            Case 4 'credit rating of E
                dAdjRate = dIntRate * 1.25
        End Select

        'sets the number of payments depending on the payment plan
        If t1Years.Checked Then

            dPlanTier = 36

        ElseIf t2Years.Checked Then

            dPlanTier = 60

        ElseIf t3Years.Checked Then

            dPlanTier = 84

        End If

        '=====Error Cases=====

        'what happens when both a credit score and payment plan are not selected
        If crdRate.SelectedIndex = -1 And t1Years.Checked = False And t2Years.Checked = False And t3Years.Checked = False Then

            MsgBox("Please choose a credit score from the list and select a payment plan.", MsgBoxStyle.OKOnly, "Error")
            errorMsg = True

            'what happens when a credit score is not selected or input incorrectly
        ElseIf crdRate.SelectedIndex = -1 Then

            MsgBox("Please choose a credit score from the list.", MsgBoxStyle.OKOnly, "Error")
            errorMsg = True

            'what happens if a payment plan is not selected
        ElseIf t1Years.Checked = False And t2Years.Checked = False And t3Years.Checked = False Then

            MsgBox("Please select a payment plan.", MsgBoxStyle.OKOnly, "Error")
            errorMsg = True

        End If

        '==========Calculation==========

        'calculates and outputs the monthly pay if there is no error in the form
        If errorMsg = False Then

            dPayment = Pmt(dAdjRate, dPlanTier, -lnAmount.Text)
            outMonthPay.Text = Format(dPayment, "Currency")

        End If

    End Sub
End Class
