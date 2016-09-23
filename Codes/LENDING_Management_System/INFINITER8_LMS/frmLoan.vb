
Imports MySql.Data.MySqlClient
Public Class frmLoan
    Public LSearch As Boolean
    Dim adding As Boolean
    Dim updatin As Boolean


    Private Sub GetLoanNo()
        Try
            sqL = "SELECT LoanNo FROM loan Order By LoanNo Desc"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                lblLoanNo.Text = Val(dr(0)) + 1
            Else
                lblLoanNo.Text = 22001000

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub LoadTotalRemittance()
     
        Try
            sqL = "SELECT SUM(TotalAmount) as TotalRemitted FROM remittance WHERE LoanNo ='" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
               
                txtTotalRem.Text = Format(dr(0), "#,##0.00")
            End If
        Catch ex As Exception
            '   MsgBox(ex.Message)
            txtTotalRem.Text = ""
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub AddLoan()
        Dim interestAmount As Double
        Dim matAmount As Double
        Dim weeklyAmount As Double
        Dim pAmount As Double
        Dim EarnedInt As Double

        EarnedInt = 0.0

        interestAmount = txtInterestAmount.Text.Replace(",", "")
        matAmount = txtMatValue.Text.Replace(",", "")
        weeklyAmount = txtWeeklyRem.Text.Replace(",", "")
        pAmount = txtPricipalLoan.Text.Replace(",", "")

        Try
            sqL = "INSERT INTO loan VALUES('" & lblLoanNo.Text & "', '" & txtDuration.Text & "', '" & txtEffectiveDate.Text & "', " & pAmount & "," & Val(cmbInterestRate.Text) & ", " & weeklyAmount & ", '" & txtMatDate.Text & "', " & matAmount & ", " & pAmount & ", " & interestAmount & ", " & matAmount & ", " & interestAmount & ", '" & txtBorrowerNo.Text & "', " & EarnedInt & ", '" & cmbTerm.Text & "','')"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                ' MsgBox("Loan successfully added.", MsgBoxStyle.Information, "Add Load")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub UpdateLoan()
        Dim interestAmount As Double
        Dim matAmount As Double
        Dim weeklyAmount As Double
        Dim pAmount As Double
        Dim unInterest As Double
        Dim bAmount As Double

        interestAmount = txtInterestAmount.Text.Replace(",", "")
        matAmount = txtMatValue.Text.Replace(",", "")
        weeklyAmount = txtWeeklyRem.Text.Replace(",", "")
        pAmount = txtPricipalLoan.Text.Replace(",", "")
        unInterest = txtUnEarnedInterest.Text.Replace(",", "")
        bAmount = txtBalance.Text.Replace(",", "")

        Try
            sqL = "UPDATE loan SET BorrowerNo = '" & txtBorrowerNo.Text & "', Duration = " & txtDuration.Text & ", EffectiveDate = '" & txtEffectiveDate.Text & "', PrincipalAmount = " & pAmount & ", InterestRate = " & Val(cmbInterestRate.Text) & ", WRemittance = " & weeklyAmount & ", MatDate = '" & txtMatDate.Text & "', MatValue = " & matAmount & ", OutstandingCap = " & pAmount & ", UnearnedInterest = " & unInterest & ", Balance = " & bAmount & ", InterestAmount = " & interestAmount & ", PaymentTerm = '" & cmbTerm.Text & "' WHERE LoanNo = '" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Changes successfully saved..", MsgBoxStyle.Information, "Update Loan")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
        
    End Sub

    Private Sub AddLedger()
        Dim dateDue As String
        '  Dim mAmount As Double
        Dim aReturn As Double
        Dim aInterest As Double
        Dim totalAmount As Double


        Dim i As Integer

        For i = 0 To dgw.Rows.Count - 1
            dateDue = dgw.Rows(i).Cells(0).Value
            totalAmount = dgw.Rows(i).Cells(1).Value
            aReturn = dgw.Rows(i).Cells(2).Value
            aInterest = dgw.Rows(i).Cells(3).Value
            ' totalAmount = dgw.Rows(i).Cells(4).Value

            Try
                sqL = "INSERT INTO ledger(DueDate, TotalAmount, ReturnAmount, Interest, LoanNo, isRemitted) VALUES('" & dateDue & "', " & totalAmount & ", " & aReturn & ", " & aInterest & ", '" & lblLoanNo.Text & "','NO')"
                ConnDB()
                cmd = New MySqlCommand(sqL, conn)
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                cmd.Dispose()
                conn.Close()
            End Try
        Next
        MsgBox("Loan successfully added.", MsgBoxStyle.Information, "Add Load")
    End Sub

    Private Sub DeleteLedger()
        Try
            sqL = "DELETE FROM ledger WHERE LoanNO = '" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub



    Private Sub LoadIntRate()
        Try
            sqL = "SELECT IntRate FROm interestrate ORDER BY IntRate"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            cmbInterestRate.Items.Clear()
            Do While dr.Read = True
                cmbInterestRate.Items.Add(dr(0))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub GetLoanInfo()
        Try
            sqL = "SELECT BorrowerNo, Duration, EffectiveDate, InterestRate, PrincipalAmount, InterestAmount, MatValue, MatDate, WRemittance, OutstandingCap, UnEarnedInterest, Balance, PaymentTerm FROM loan WHERE LoanNo = '" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                txtBorrowerNo.Text = dr(0)
                txtDuration.Text = dr(1)
                txtEffectiveDate.Text = dr(2)
                cmbInterestRate.Text = dr(3)
                txtPricipalLoan.Text = Format(dr(4), "#,##0.00")
                txtInterestAmount.Text = Format(dr(5), "#,##0.00")
                txtMatValue.Text = Format(dr(6), "#,##0.00")
                txtMatDate.Text = dr(7)
                txtWeeklyRem.Text = Format(dr(8), "#,##0.00")
                txtOutsCapital.Text = Format(dr(9), "#,##0.00")
                txtUnEarnedInterest.Text = Format(dr(10), "#,##0.00")
                txtBalance.Text = Format(dr(11), "#,##0.00")
                cmbTerm.Text = dr(12)
            Else
                ' MsgBox("Not found", MsgBoxStyle.Critical, "Search Record")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()

        End Try
    End Sub

    Private Sub LoadLedger()
        Dim Rbal As String
        Dim pAmount As String
        Try
            sqL = "SELECT DueDate, TotalAmount, ReturnAmount, Interest, PaidDate,  RemainingBalance, PaidAmount  FROM ledger WHERE LoanNo = '" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                If dr(5) = 0 And dr(6) = 0 Then
                    Rbal = ""
                    pAmount = ""
                   
                Else
                    Rbal = Format(dr(5), "#,##0.00").ToString()
                    pAmount = Format(dr(6), "#,##0.00").ToString()
                End If
                dgw.Rows.Add(dr(0).ToString(), Format(dr(1), "#,##0.00"), Format(dr(2), "#,##0.00"), Format(dr(3), "#,##0.00"), dr(4).ToString(), pAmount, Rbal)
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Function valUpdate() As Boolean
        Dim retBol As Boolean
        Try
            sqL = "SELECT * FROM loan WHERE MatAmount > Balance AND LoanNo ='" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                MsgBox("Changes cannot be done because this loan is on going already..", MsgBoxStyle.Critical, "Error")
                retBol = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try

        Return retBol
    End Function

    Private Sub GetBorrowerName()
        Try
            sqL = "SELECT CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName FROM loan as L, borrower as B WHERE L.BorrowerNo = B. BorrowerNo AND LoanNo = '" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                txtName.Text = dr(0)
            Else
                txtName.Text = ""
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub ComputeLoan()
        'dito sa procedure nato ka magseset kung monhty, weekly, daily
        'maglagay ka sa form ng additional function (combo box or checkbox) in which the user can select if it is monthly, weekly or daily

        Dim interest As Double
        Dim princiapalLoan As Double
        Dim matValue As Double


        Dim weeklyRem As Double
        Dim dailyRem As Double
        Dim monthlyRem As Double

        Dim effDate As Date
        Dim dueDate As Date
        Dim MatDate As Date

        Dim i As Integer
        Dim weekInt As Integer
        Dim dailyInt As Integer
        Dim durations As Integer = Val(txtDuration.Text)

        'gumamit ka ng if else statement
        'like for example if the user select monthly

        'if combobox.text ='Monthly'
        '       lagay mo dito lahat ng computation ng monthly (interest, maturity value, maturity date, monthly remittance,)
        '       (itong block nato ang ma eexecute)
        'elseif combobox.text='Weekly'
        '       lagay mo dito lahat ng computation ng weekly (interest, maturity value, maturity date, weekly remittance,)
        'elseif combobox.text='Daily'
        '       lagay mo dito lahat ng computation ng Daily (interest, maturity value, maturity date, Daily remittance,)

        If cmbTerm.Text = "Monthly" Then
            princiapalLoan = txtPricipalLoan.Text.Replace(",", "")
            'Interest
            interest = princiapalLoan * (Val(cmbInterestRate.Text) / 100)
            txtInterestAmount.Text = Format(interest, "#,##0.00")
            'Maturity Value
            matValue = princiapalLoan + interest
            txtMatValue.Text = Format(matValue, "#,##0.00")
            'Maturity Date

            effDate = Date.ParseExact(txtEffectiveDate.Text, "MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            MatDate = effDate.AddMonths(1 * Val(txtDuration.Text.Replace(",", "")))
            txtMatDate.Text = MatDate.ToString("MM/dd/yyyy")
            'Weekly Remittance
            monthlyRem = matValue / Val(txtDuration.Text)
            txtWeeklyRem.Text = Format(monthlyRem, "#,##0.00")

            ' Outstanding Capital, Unearned Interest, Balance
            txtOutsCapital.Text = Format(princiapalLoan, "#,##0.00")
            txtUnEarnedInterest.Text = txtInterestAmount.Text
            txtBalance.Text = txtMatValue.Text

            'Get the Ledger Value
            dgw.Rows.Clear()
            Do While i < durations
                i += 1
                'also here set an if else statement
                'if monthly set mo ang value ng weekInt = 30, if weekly = 7, if daily=1
                weekInt += 1

                dueDate = effDate.AddMonths(weekInt)
                dgw.Rows.Add(dueDate.ToString("MM/dd/yyyy"), Format((princiapalLoan / durations) + (interest / durations), "#,##0.00"), Format((princiapalLoan / durations), "#,##0.00"), Format((interest / durations), "#,##0.00"), "")
            Loop
        ElseIf cmbTerm.Text = "Weekly" Then
            princiapalLoan = txtPricipalLoan.Text.Replace(",", "")
            'Interest
            interest = princiapalLoan * (Val(cmbInterestRate.Text) / 100)
            txtInterestAmount.Text = Format(interest, "#,##0.00")
            'Maturity Value
            matValue = princiapalLoan + interest
            txtMatValue.Text = Format(matValue, "#,##0.00")
            'Maturity Date

            effDate = Date.ParseExact(txtEffectiveDate.Text, "MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            MatDate = effDate.AddDays(7 * Val(txtDuration.Text.Replace(",", "")))
            txtMatDate.Text = MatDate.ToString("MM/dd/yyyy")
            'Weekly Remittance
            weeklyRem = matValue / Val(txtDuration.Text)
            txtWeeklyRem.Text = Format(weeklyRem, "#,##0.00")

            ' Outstanding Capital, Unearned Interest, Balance
            txtOutsCapital.Text = Format(princiapalLoan, "#,##0.00")
            txtUnEarnedInterest.Text = txtInterestAmount.Text
            txtBalance.Text = txtMatValue.Text

            'Get the Ledger Value
            dgw.Rows.Clear()
            Do While i < durations
                i += 1
                'also here set an if else statement
                'if monthly set mo ang value ng weekInt = 30, if weekly = 7, if daily=1
                weekInt += 7

                dueDate = effDate.AddDays(weekInt)
                dgw.Rows.Add(dueDate.ToString("MM/dd/yyyy"), Format((princiapalLoan / durations) + (interest / durations), "#,##0.00"), Format((princiapalLoan / durations), "#,##0.00"), Format((interest / durations), "#,##0.00"), "")
            Loop

        ElseIf cmbTerm.Text = "Daily" Then
            princiapalLoan = txtPricipalLoan.Text.Replace(",", "")
            'Interest
            interest = princiapalLoan * (Val(cmbInterestRate.Text) / 100)
            txtInterestAmount.Text = Format(interest, "#,##0.00")
            'Maturity Value
            matValue = princiapalLoan + interest
            txtMatValue.Text = Format(matValue, "#,##0.00")
            'Maturity Date

            effDate = Date.ParseExact(txtEffectiveDate.Text, "MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            MatDate = effDate.AddDays(Val(txtDuration.Text.Replace(",", "")))
            txtMatDate.Text = MatDate.ToString("MM/dd/yyyy")
            'Weekly Remittance
            dailyRem = matValue / Val(txtDuration.Text)
            txtWeeklyRem.Text = Format(dailyRem, "#,##0.00")

            ' Outstanding Capital, Unearned Interest, Balance
            txtOutsCapital.Text = Format(princiapalLoan, "#,##0.00")
            txtUnEarnedInterest.Text = txtInterestAmount.Text
            txtBalance.Text = txtMatValue.Text

            'Get the Ledger Value
            dgw.Rows.Clear()
            Do While i < durations
                i += 1
                'also here set an if else statement
                'if monthly set mo ang value ng weekInt = 30, if weekly = 7, if daily=1
                dailyInt += 1

                dueDate = effDate.AddDays(dailyInt)
                dgw.Rows.Add(dueDate.ToString("MM/dd/yyyy"), Format((princiapalLoan / durations) + (interest / durations), "#,##0.00"), Format((princiapalLoan / durations), "#,##0.00"), Format((interest / durations), "#,##0.00"), "")
            Loop

        ElseIf cmbTerm.Text = "Semi Monthly" Then
            princiapalLoan = txtPricipalLoan.Text.Replace(",", "")
            'Interest
            interest = princiapalLoan * (Val(cmbInterestRate.Text) / 100)
            txtInterestAmount.Text = Format(interest, "#,##0.00")
            'Maturity Value
            matValue = princiapalLoan + interest
            txtMatValue.Text = Format(matValue, "#,##0.00")
            'Maturity Date

            effDate = Date.ParseExact(txtEffectiveDate.Text, "MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            MatDate = effDate.AddDays(7 * Val(txtDuration.Text.Replace(",", "")))
            txtMatDate.Text = MatDate.ToString("MM/dd/yyyy")
            'Weekly Remittance
            weeklyRem = matValue / Val(txtDuration.Text)
            txtWeeklyRem.Text = Format(weeklyRem, "#,##0.00")

            ' Outstanding Capital, Unearned Interest, Balance
            txtOutsCapital.Text = Format(princiapalLoan, "#,##0.00")
            txtUnEarnedInterest.Text = txtInterestAmount.Text
            txtBalance.Text = txtMatValue.Text
            Dim dblNo As Double
            'Get the Ledger Value
            dgw.Rows.Clear()
            Do While i < durations
                i += 1
                'also here set an if else statement
                'if monthly set mo ang value ng weekInt = 30, if weekly = 7, if daily=1
                dblNo += 15

                dueDate = effDate.AddDays(dblNo)
                dgw.Rows.Add(dueDate.ToString("MM/dd/yyyy"), Format((princiapalLoan / durations) + (interest / durations), "#,##0.00"), Format((princiapalLoan / durations), "#,##0.00"), Format((interest / durations), "#,##0.00"), "")
            Loop

        End If

    End Sub

    Private Sub Cleafields()
        txtBorrowerNo.Text = ""
        txtDuration.Text = ""
        txtEffectiveDate.Text = ""
        cmbInterestRate.SelectedIndex = -1
        txtPricipalLoan.Text = ""
        txtInterestAmount.Text = ""
        txtMatValue.Text = ""
        txtMatDate.Text = ""
        txtWeeklyRem.Text = ""
        txtTotalRem.Text = ""
        txtOutsCapital.Text = ""
        txtUnEarnedInterest.Text = ""
        txtBalance.Text = ""
        lblLoanNo.Text = ""
        txtName.Text = ""
        cmbTerm.SelectedIndex = -1
        dgw.Rows.Clear()
    End Sub

    Private Sub EnabledFields()
        txtBorrowerNo.Enabled = True
        txtDuration.Enabled = True
        txtEffectiveDate.Enabled = True
        cmbInterestRate.Enabled = True
        txtPricipalLoan.Enabled = True
        cmbTerm.Enabled = True

    End Sub

    Private Sub DisabledFields()
        txtBorrowerNo.Enabled = False
        txtDuration.Enabled = False
        txtEffectiveDate.Enabled = False
        cmbInterestRate.Enabled = False
        txtPricipalLoan.Enabled = False
        cmbTerm.Enabled = False
    End Sub


    Private Sub EnabledButtons()
        btnAdd.Enabled = True
        btnUpdate.Enabled = True
        btnSearch.Enabled = True
        btnClose.Enabled = True

        Button1.Enabled = False
        btnSave.Enabled = False
        btnCancel.Enabled = False
    End Sub

    Private Sub DisabledButtons()
        btnAdd.Enabled = False
        btnUpdate.Enabled = False
        btnSearch.Enabled = False
        btnClose.Enabled = False

        Button1.Enabled = True
        btnSave.Enabled = True
        btnCancel.Enabled = True
    End Sub

    Private Function IsAllFieldsFilled() As Boolean
        Dim bolRet As Boolean
        Try
            If cmbTerm.Text = "" Then
                bolRet = False
            ElseIf txtDuration.Text = "" Then
                bolRet = False
            ElseIf cmbInterestRate.Text = "" Then
                bolRet = False
            ElseIf txtEffectiveDate.Text.Length < 6 Then
                bolRet = False
            ElseIf txtPricipalLoan.Text = "" Then
                bolRet = False
            Else
                bolRet = True
            End If
        Catch ex As Exception

        End Try
        Return bolRet
    End Function

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtPricipalLoan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPricipalLoan.KeyPress
        If e.KeyChar = ControlChars.Cr Then
            If IsAllFieldsFilled() = True Then
                ComputeLoan()
            Else
                MsgBox("Please fill all fields.", MsgBoxStyle.Exclamation, "Warning")
            End If

        End If
    End Sub


    Private Sub frmLoan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DisabledFields()
        EnabledButtons()
        Cleafields()
        LoadIntRate()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LSearch = True
        frmLoadBorrower.ShowDialog()
    End Sub


    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        adding = True
        updatin = False

        EnabledFields()
        DisabledButtons()
        Cleafields()
        GetLoanNo()
        txtBorrowerNo.Focus()
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If lblLoanNo.Text = "" Then
            MsgBox("Please select record to change..", MsgBoxStyle.Exclamation, "Update Loan")
            Exit Sub
        End If

        Dim pA As Double = txtMatValue.Text.Replace(",", "")
        Dim bA As Double = txtBalance.Text.Replace(",", "")
        If pA > bA Then
            MsgBox("Loan cannot be updated, because it is already on going..", MsgBoxStyle.Exclamation, "Warning")
            Exit Sub
        End If


        adding = False
        updatin = True
        EnabledFields()
        DisabledButtons()
        txtPricipalLoan.Focus()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        LSearch = True
        frmLoadLoan.ShowDialog()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If adding = True Then
            AddLoan()
            AddLedger()
        Else
            UpdateLoan()
            DeleteLedger()
            AddLedger()
        End If

        Cleafields()
        DisabledFields()
        EnabledButtons()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Cleafields()
        DisabledFields()
        EnabledButtons()
    End Sub


    Private Sub lblLoanNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLoanNo.TextChanged

        If LSearch = True Then
            GetLoanInfo()
            LoadLedger()
            GetBorrowerName()
            LoadTotalRemittance()
            LSearch = False
        End If
       
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If lblLoanNo.Text = "" Then
            MsgBox("Please select record from LOAN..", MsgBoxStyle.Exclamation, "Warning")
            Exit Sub
        End If
        frmRemitted.ShowDialog()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If lblLoanNo.Text = "" Then
            MsgBox("Please select record to print..", MsgBoxStyle.Exclamation, "Print")
            Exit Sub
        End If
        'frmPrintLedger.ShowDialog()
        frmReportLedger.ShowDialog()
    End Sub

    Private Sub txtEffectiveDate_Leave(sender As System.Object, e As System.EventArgs) Handles txtEffectiveDate.Leave
        Try
            Dim BDDate As Date = Date.ParseExact(txtEffectiveDate.Text, "MM/dd/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            If IsDate(BDDate) Then

            Else
                MsgBox("Invalid Date Format. Please enter a valid date.", MsgBoxStyle.Exclamation, "Validation")
                txtEffectiveDate.Focus()
            End If
        Catch ex As Exception
            MsgBox("Invalid Date Format. Please enter a valid date.", MsgBoxStyle.Exclamation, "Validation")
            txtEffectiveDate.Focus()
        End Try
    End Sub

   
End Class