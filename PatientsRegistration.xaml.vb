Imports System.Data
Imports System.Data.SqlClient


Public Class PatientsRegistration
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        txtidno.Text = ""
        txtname.Text = ""
        txtaddress.Text = ""
        txtmarital.Text = ""
        lblgender.Content = ""
        txtregamount.Text = "0.00"
        'tryn add new no
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

        cmd.CommandText = "Select * from patients_registration"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtidno.Text = 0
                Else
                    newNum = 1 + last
                    txtidno.Text = newNum
                End If
            Catch ex As Exception

            End Try
        End With

    End Sub

    Private Sub optfemale_Click(sender As Object, e As RoutedEventArgs) Handles optfemale.Click
        lblgender.Content = "Female"
    End Sub

    Private Sub optmale_Click(sender As Object, e As RoutedEventArgs) Handles optmale.Click
        lblgender.Content = "male"
    End Sub

    Private Sub buttonSave_Click(sender As Object, e As RoutedEventArgs) Handles buttonSave.Click
        Dim connectionstring As String
        'getting the connection string from the Main Navigation form
        Dim mainConn As New MainNavigation
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "insert into patients_registration(Id_no,name,Address,Marital,Sex,Date,Registration_Amount) Values
('" & txtidno.Text & "','" & txtname.Text & "','" & txtaddress.Text & "','" & txtmarital.Text & "','" & lblgender.Content & "'
,'" & dtpicker.Text & "','" & txtregamount.Text & "')"

            cmd.ExecuteNonQuery()
            MessageBox.Show("Patient Registered Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception

        End Try

        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter ID. No to search")

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

                txtname.Text = .GetValue(0)
                txtaddress.Text = .GetValue(1)
                txtmarital.Text = .GetValue(2)
                lblgender.Content = .GetValue(3)
                dtpicker.Text = .GetValue(4)
                txtidno.Text = .GetValue(5)
                txtregamount.Text = .GetValue(6)
            End With
            dr.Close()
        Catch ex As Exception

        End Try


    End Sub

    Private Sub buttonUpdate_Click(sender As Object, e As RoutedEventArgs) Handles buttonUpdate.Click
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "update patients_registration set Id_no='" & txtidno.Text & "',name='" & txtname.Text & "',Address='" & txtaddress.Text & "',Marital='" & txtmarital.Text & "',
        Sex='" & lblgender.Content & "',Date='" & dtpicker.Text & "',Registration_Amount='" & txtregamount.Text & "'"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Patient Updates Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Patients ID to Delete")

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
            Try
                With dr
                    .Read()

                    txtname.Text = .GetValue(0)
                    txtaddress.Text = .GetValue(1)
                    txtmarital.Text = .GetValue(2)
                    lblgender.Content = .GetValue(3)
                    dtpicker.Text = .GetValue(4)
                    txtidno.Text = .GetValue(5)
                    txtregamount.Text = .GetValue(6)
                End With
                dr.Close()
            Catch ex As Exception

            End Try

            Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
                Case vbYes
                    ''Deleting current record
                    cmd.CommandText = "Delete from patients_registration where Id_no='" & Val(a) & "'"
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("its been a pleasure Freeing up some space, Delete Succesfull")
                    cs.Close()
                Case vbNo
                    Exit Sub
            End Select
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

        cmd.CommandText = "Select * from patients_registration"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(Table)
        With Table
            Try

                txtname.Text = .Rows(position)(0).ToString
                txtaddress.Text = .Rows(position)(1).ToString
                txtmarital.Text = .Rows(position)(2).ToString
                lblgender.Content = .Rows(position)(3).ToString
                dtpicker.Text = .Rows(position)(4).ToString
                txtidno.Text = .Rows(position)(5).ToString
                txtregamount.Text = .Rows(position)(6).ToString
            Catch ex As Exception

            End Try
        End With
    End Sub

    Private Sub btnmovenext_Click(sender As Object, e As RoutedEventArgs) Handles btnmovenext.Click
        Index += 1
        showData(Index)
    End Sub

    Private Sub btnprevious_Click(sender As Object, e As RoutedEventArgs) Handles btnprevious.Click
        Index -= 1
        showData(Index)
    End Sub

    Private Sub btnfirst_Click(sender As Object, e As RoutedEventArgs) Handles btnfirst.Click
        Index = 0
        showData(Index)
    End Sub

    Private Sub btnmovelast_Click(sender As Object, e As RoutedEventArgs) Handles btnmovelast.Click
        Index = Table.Rows.Count() - 1
        showData(Index)
    End Sub

    Private Sub btndisplay_Click(sender As Object, e As RoutedEventArgs) Handles btndisplay.Click
        Dim report As New ReportViewer
        report.Show()
        Dim myrep As New Ripoti_registration

        myrep.Load("..\\Ripoti_registration.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub

    Private Sub PatientsRegistration_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            index = table.Rows.Count()
            showData(index)
            txtmarital.Items.Add("Single")
            txtmarital.Items.Add("Married")

        Catch ex As Exception

        End Try

    End Sub
End Class
