Imports System.Data
Imports System.Data.SqlClient
Imports System.Printing
Imports System.Windows.Documents
Public Class BillingDetails
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub optadmitted_Click(sender As Object, e As RoutedEventArgs) Handles optadmitted.Click
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Ward_Admission"
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cboadmissionNo.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub optnotadmitted_Click(sender As Object, e As RoutedEventArgs) Handles optnotadmitted.Click
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Patients_Diagnosis where AdmissionCode='" & "0" & "'"
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cboDiagnosisNo.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()
        Catch ex As Exception

        End Try
    End Sub


    Private Sub cboadmissionNo_MouseLeave(sender As Object, e As MouseEventArgs) Handles cboadmissionNo.MouseLeave
        Try
            Dim a As Integer
            a = cboadmissionNo.Text

            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader

            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Ward_Admission where Ward_admission_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader

            With dr
                .Read()
                txtPatientsID.Text = .GetValue(2)
                txtname.Text = .GetValue(3)
                txtdocname.Text = .GetValue(5)
                txtwardno.Text = .GetValue(8)
                txtwardname.Text = .GetValue(9)
                txtBedNo.Text = .GetValue(10)
                txtadmdate.Text = .GetValue(11)
                txthospitalbill.Text = .GetValue(12)
            End With
            dr.Close()

        Catch ex As Exception
        End Try

    End Sub

    Private Sub txtPatientsID_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtPatientsID.TextChanged
        Try
            Dim a As Integer
            a = txtPatientsID.Text

            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader
            Try
                cmd.Connection = cs
                cs.Open()

                cmd.CommandText = "Select * from patients_registration where Id_no='" & Val(a) & "'"
                dr = cmd.ExecuteReader

                With dr
                    .Read()
                    txtadmbills.Text = .GetValue(6)
                End With
                dr.Close()
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboDiagnosisNo_MouseLeave(sender As Object, e As MouseEventArgs) Handles cboDiagnosisNo.MouseLeave
        Try
            Dim a As Integer
            a = cboDiagnosisNo.Text

            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader

            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Patients_Diagnosis where Diagnosis_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader
            Try
                With dr
                    .Read()

                    txtPatientsID.Text = .GetValue(2)
                    txtname.Text = .GetValue(3)
                    txtdocname.Text = .GetValue(5)

                End With
                dr.Close()
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click

        txtbillno.Text = ""
        cboadmissionNo.Text = ""
        cboDiagnosisNo.Text = ""
        txtwardno.Text = ""
        txtPatientsID.Text = ""
        txtwardname.Text = ""
        txtBedNo.Text = ""
        txtname.Text = ""
        txtdocname.Text = ""
        txtadmdate.Text = ""
        dtpicker.Text = ""
        txtadmbills.Text = ""
        txthospitalbill.Text = ""
        txtamountpaid.Text = ""
        txtbalance.Text = ""


        'add new num
        index = table.Rows.Count() - 1
        AddIndex(index)
    End Sub
    Public Sub AddIndex(position As Integer)

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)

        Dim cmd, cmd2 As New SqlCommand
        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Bills"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtbillno.Text = 0
                Else
                    newNum = 1 + last
                    txtbillno.Text = newNum
                End If
            Catch ex As Exception
            End Try
        End With

    End Sub
    Private Sub buttonSave1_Click_1(sender As Object, e As RoutedEventArgs) Handles buttonSave1.Click
        Dim connectionstring As String
        'getting the connection string from the Main Navigation form
        Dim mainConn As New MainNavigation
        connectionstring = lblstring.Content
        Dim diag As Integer

        diag = 0
        diag = Val(cboDiagnosisNo.Text)
        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "insert into Bills(BillNo,AdmissionNo,DiagnosisNo,WardNo,PatientsID,WardName,BedNo,Patients_Name,Doc_Name,AdmDate,Date,AdmissionBill,HospitalBill,AmountPaud,Balance) Values
        ('" & txtbillno.Text & "','" & cboadmissionNo.Text & "','" & diag & "','" & txtwardno.Text & "','" & txtPatientsID.Text & "','" & txtwardname.Text & "',
        '" & txtBedNo.Text & "','" & txtname.Text & "','" & txtdocname.Text & "','" & txtadmdate.Text & "','" & dtpicker.Text & "','" & txtadmbills.Text & "','" & txthospitalbill.Text & "','" & txtamountpaid.Text & "','" & txtbalance.Text & "')"

            cmd.ExecuteNonQuery()
            MessageBox.Show("Bills Updated", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Bill No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Bills where BillNo='" & Val(a) & "'"
        dr = cmd.ExecuteReader

        With dr
            Try
                .Read()
                txtbillno.Text = .GetValue(0)
                cboadmissionNo.Text = .GetValue(1)
                cboDiagnosisNo.Text = .GetValue(2)
                txtwardno.Text = .GetValue(3)
                txtPatientsID.Text = .GetValue(4)
                txtwardname.Text = .GetValue(5)
                txtBedNo.Text = .GetValue(6)
                txtname.Text = .GetValue(7)
                txtdocname.Text = .GetValue(8)
                txtadmdate.Text = .GetValue(9)
                dtpicker.Text = .GetValue(10)
                txtadmbills.Text = .GetValue(11)
                txthospitalbill.Text = .GetValue(12)
                txtamountpaid.Text = .GetValue(13)
                txtbalance.Text = .GetValue(14)
            Catch ex As Exception

            End Try
        End With
        dr.Close()

    End Sub

    Private Sub buttonUpdate_Click(sender As Object, e As RoutedEventArgs) Handles buttonUpdate.Click
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        If cboadmissionNo.Text = "" Then
            Try
                cmd.Connection = cs
                cs.Open()
                cmd.CommandText = "update Bills set BillNo='" & txtbillno.Text & "',AdmissionNo='" & "0" & "',DiagnosisNo='" & cboDiagnosisNo.Text & "',WardNo='" & txtwardno.Text & "',PatientsID='" & txtPatientsID.Text & "'
			,WardName='" & txtwardname.Text & "',BedNo='" & txtBedNo.Text & "',Patients_Name='" & txtname.Text & "',Doc_Name='" & txtdocname.Text & "'
			,AdmDate='" & txtadmdate.Text & "',Date='" & dtpicker.Text & "',AdmissionBill='" & txtadmbills.Text & "',HospitalBill='" & txthospitalbill.Text & "',AmountPaud='" & txtamountpaid.Text & "',Balance='" & txtbalance.Text & "'"

                cmd.ExecuteNonQuery()
                MessageBox.Show("Bill Updated Succesful", "Kapenguria Patients MIS", vbOK)
                cs.Close()
            Catch ex As Exception

            End Try
        End If

        If cboDiagnosisNo.Text = "" Then
            Try
                cmd.Connection = cs
                cs.Open()
                cmd.CommandText = "update Bills set BillNo='" & txtbillno.Text & "',AdmissionNo='" & cboadmissionNo.Text & "',DiagnosisNo='" & "0" & "',WardNo='" & txtwardno.Text & "',PatientsID='" & txtPatientsID.Text & "'
			,WardName='" & txtwardname.Text & "',BedNo='" & txtBedNo.Text & "',Patients_Name='" & txtname.Text & "',Doc_Name='" & txtdocname.Text & "'
			,AdmDate='" & txtadmdate.Text & "',Date='" & dtpicker.Text & "',AdmissionBill='" & txtadmbills.Text & "',HospitalBill='" & txthospitalbill.Text & "',AmountPaud='" & txtamountpaid.Text & "',Balance='" & txtbalance.Text & "'"

                cmd.ExecuteNonQuery()
                MessageBox.Show("Bill Updated Succesful", "Kapenguria Patients MIS", vbOK)
                cs.Close()
            Catch ex As Exception

            End Try
        End If
        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Bill No Delete")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Bills where BillNo='" & Val(a) & "'"
        dr = cmd.ExecuteReader
        Try
            With dr
                .Read()

                txtbillno.Text = .GetValue(0)
                cboadmissionNo.Text = .GetValue(1)
                cboDiagnosisNo.Text = .GetValue(2)
                txtwardno.Text = .GetValue(3)
                txtPatientsID.Text = .GetValue(4)
                txtwardname.Text = .GetValue(5)
                txtBedNo.Text = .GetValue(6)
                txtname.Text = .GetValue(7)
                txtdocname.Text = .GetValue(8)
                txtadmdate.Text = .GetValue(9)
                dtpicker.Text = .GetValue(10)
                txtadmbills.Text = .GetValue(11)
                txthospitalbill.Text = .GetValue(12)
                txtamountpaid.Text = .GetValue(13)
                txtbalance.Text = .GetValue(14)
            End With
            dr.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Select Case MessageBox.Show("Please confirm you want to delete this record", "Freeing Up Some Space", vbYesNo)
            Case vbYes
                ''Deleting current record
                cmd.CommandText = "Delete from Bills where BillNo='" & Val(a) & "'"
                cmd.ExecuteNonQuery()
                MessageBox.Show("its been a pleasure Freeing up some space, Delete Succesfull")
                cs.Close()
            Case vbNo
                Exit Sub
        End Select

    End Sub

    Private Sub txtadmbills_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtadmbills.TextChanged
        Try

            Dim total, admbill, hosbill, paid As Long
            admbill = Val(txtadmbills.Text)
            hosbill = Val(txthospitalbill.Text)
            paid = Val(txtamountpaid.Text)
            total = hosbill + admbill

            txttotal.Text = total
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txthospitalbill_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txthospitalbill.TextChanged
        Try

            Dim total, admbill, hosbill As Long
            admbill = Val(txtadmbills.Text)
            hosbill = Val(txthospitalbill.Text)

            total = hosbill + admbill

            txttotal.Text = total
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txttotal_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txttotal.TextChanged
        Try

            Dim total, balance, paid As Long

            paid = Val(txtamountpaid.Text)
            total = Val(txttotal.Text)
            balance = total - paid
            txtbalance.Text = balance
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtamountpaid_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtamountpaid.TextChanged
        Try

            Dim total, balance, paid As Long

            paid = Val(txtamountpaid.Text)
            total = Val(txttotal.Text)
            balance = total - paid
            txtbalance.Text = balance
        Catch ex As Exception
        End Try
    End Sub
    Public Sub showData(position As Integer)
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)

        Dim cmd, cmd2 As New SqlCommand
        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Bills"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(Table)
        With Table
            Try
                txtbillno.Text = .Rows(position)(0).ToString
                cboadmissionNo.Text = .Rows(position)(1).ToString
                cboDiagnosisNo.Text = .Rows(position)(2).ToString
                txtwardno.Text = .Rows(position)(3).ToString
                txtPatientsID.Text = .Rows(position)(4).ToString
                txtwardname.Text = .Rows(position)(5).ToString
                txtBedNo.Text = .Rows(position)(6).ToString
                txtname.Text = .Rows(position)(7).ToString
                txtdocname.Text = .Rows(position)(8).ToString
                txtadmdate.Text = .Rows(position)(9).ToString
                dtpicker.Text = .Rows(position)(10).ToString
                txtadmbills.Text = .Rows(position)(11).ToString
                txthospitalbill.Text = .Rows(position)(12).ToString
                txtamountpaid.Text = .Rows(position)(13).ToString
                txtbalance.Text = .Rows(position)(14).ToString
            Catch ex As Exception

            End Try
        End With
    End Sub

    Private Sub btnnext_Click(sender As Object, e As RoutedEventArgs) Handles btnnext.Click
        index += 1
        showData(index)
    End Sub

    Private Sub btnprevious1_Click(sender As Object, e As RoutedEventArgs) Handles btnprevious1.Click
        index -= 1
        showData(index)
    End Sub

    Private Sub btnfirst1_Click(sender As Object, e As RoutedEventArgs) Handles btnfirst1.Click
        index = 0
        showData(index)
    End Sub

    Private Sub btnlast_Click(sender As Object, e As RoutedEventArgs) Handles btnlast.Click
        index = Table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub btndisplay_Click(sender As Object, e As RoutedEventArgs) Handles btndisplay.Click
        Dim report As New ReportViewer
        report.Show()
        Dim myrep As New Ripoti_Billing

        myrep.Load("..\\Ripoti_Billing.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
        report.Title = "Billing Report"
    End Sub

    Private Sub BillingDetails_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            index = table.Rows.Count()
            showData(index)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonSave_Click(sender As Object, e As RoutedEventArgs) Handles buttonSave.Click
        Dim report As New Receipt1
        report.txtbillno.Text = txtbillno.Text
        report.lblstring.Content = lblstring.Content

        report.Show()

    End Sub

End Class
