Imports System.Data
Imports System.Data.SqlClient
Public Class Register_otherStaff
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        txtempno.Text = ""
        txtname.Text = ""
        cboduty.Text = ""
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

        cmd.CommandText = "Select * from other_staff"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtempno.Text = 0
                Else
                    newNum = (1 + last) + 10
                    txtempno.Text = newNum
                End If
            Catch ex As Exception

            End Try
        End With

    End Sub

    Private Sub buttonSave_Click(sender As Object, e As RoutedEventArgs) Handles buttonSave.Click
        Dim connectionstring As String
        'getting the connection string from the Main Navigation form
        Dim mainConn As New MainNavigation
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()
        cmd.CommandText = "insert into other_staff(EmploymentNo,Name,Duty) Values
('" & txtempno.Text & "','" & txtname.Text & "','" & cboduty.Text & "')"

        cmd.ExecuteNonQuery()
        MessageBox.Show("Staff Registered Succesful", "Kapenguria Patients MIS", vbOK)

        cs.Close()
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Employment. No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from other_staff where EmploymentNo='" & Val(a) & "'"
        dr = cmd.ExecuteReader

        With dr
            .Read()

            txtname.Text = .GetValue(1)
            cboduty.Text = .GetValue(2)
            txtempno.Text = .GetValue(0)
        End With
        dr.Close()

    End Sub

    Private Sub buttonUpdate_Click(sender As Object, e As RoutedEventArgs) Handles buttonUpdate.Click
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "update other_staff set EmploymentNo='" & txtempno.Text & "',Name='" & txtname.Text & "',Duty='" & cboduty.Text & "'"
        cmd.ExecuteNonQuery()
        MessageBox.Show("Staff Update Succesful", "Kapenguria Patients MIS", vbOK)

        cs.Close()
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Staff Employment No..to Delete")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from other_staff where EmploymentNo='" & Val(a) & "'"
        dr = cmd.ExecuteReader
        Try
            With dr
                .Read()

                txtname.Text = .GetValue(1)
                cboduty.Text = .GetValue(2)
                txtempno.Text = .GetValue(0)
            End With
            dr.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
            Case vbYes
                ''Deleting current record
                cmd.CommandText = "Delete from other_staff where EmploymentNo='" & Val(a) & "'"
                cmd.ExecuteNonQuery()
                MessageBox.Show("its been a pleasure Freeing up some space, Delete Succesfull")
                cs.Close()
            Case vbNo
                Exit Sub
        End Select

    End Sub

    Private Sub Register_otherStaff_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        cboduty.Items.Add("Management")
        cboduty.Items.Add("Receptionist")
        cboduty.Items.Add("Billing officer")

        index = table.Rows.Count()
        showData(index)
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
                txtname.Text = .Rows(position)(1).ToString
                cboduty.Text = .Rows(position)(2).ToString
                txtempno.Text = .Rows(position)(0).ToString

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
        Dim myrep As New Ripoti_staff
        report.Title = "Staff Members"
        myrep.Load("..\\Ripoti_staff.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub
End Class
