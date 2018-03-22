Imports System.Data
Imports System.Data.SqlClient

Public Class Receipt1
    Dim index As Integer = 0
    Dim table As New DataTable()


    Private Sub BillingDetails_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try

            Dim a As Integer
            a = txtbillno.Text

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
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtbillno_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtbillno.TextChanged

        Try
            Dim a As Integer
            a = txtbillno.Text

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
        Catch ex As Exception

        End Try
    End Sub

End Class
