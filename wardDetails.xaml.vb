Imports System.Data
Imports System.Data.SqlClient

Public Class wardDetails
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        txtwardno.Text = ""
        txtwardname.Text = ""
        txtbeds.Text = ""
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

        cmd.CommandText = "Select * from Wards"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtwardno.Text = 0
                Else
                    newNum = 1 + last
                    txtwardno.Text = newNum
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
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
        cmd.CommandText = "insert into Wards(Ward_No,Ward_Name,No_of_Beds) Values
('" & txtwardno.Text & "','" & txtwardname.Text & "','" & txtbeds.Text & "')"

        cmd.ExecuteNonQuery()
        MessageBox.Show("Ward Registered Succesful", "Kapenguria Patients MIS", vbOK)
        cs.Close()

        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Ward. No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Wards where Ward_No='" & Val(a) & "'"
        dr = cmd.ExecuteReader

        With dr
            .Read()
            txtwardno.Text = .GetValue(0)
            txtwardname.Text = .GetValue(2)
            txtbeds.Text = .GetValue(1)

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

        cmd.CommandText = "update Wards set Ward_No='" & txtwardno.Text & "',Ward_Name='" & txtwardname.Text & "',No_of_Beds='" & txtbeds.Text & "'"
        cmd.ExecuteNonQuery()
        MessageBox.Show("Ward Updates Succesful", "Kapenguria Patients MIS", vbOK)

        cs.Close()
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Ward No to Delete")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Wards where Ward_No='" & Val(a) & "'"
        dr = cmd.ExecuteReader
        Try
            With dr
                .Read()
                txtwardno.Text = .GetValue(0)
                txtwardname.Text = .GetValue(2)
                txtbeds.Text = .GetValue(1)
            End With
            dr.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
            Case vbYes
                ''Deleting current record
                cmd.CommandText = "Delete from Wards where Ward_No='" & Val(a) & "'"
                cmd.ExecuteNonQuery()
                MessageBox.Show("its been a pleasure Freeing up some space, Delete Succesfull")
                cs.Close()
            Case vbNo
                Exit Sub
        End Select

    End Sub
    Public Sub showData(position As Integer)
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)

        Dim cmd, cmd2 As New SqlCommand
        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Wards"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(Table)
        With Table
            Try
                txtwardno.Text = .Rows(position)(0).ToString
                txtwardname.Text = .Rows(position)(2).ToString
                txtbeds.Text = .Rows(position)(1).ToString

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End With
    End Sub

    Private Sub btnmovenext_Click(sender As Object, e As RoutedEventArgs) Handles btnmovenext.Click
        index += 1
        showData(index)
    End Sub

    Private Sub btnprevious_Click(sender As Object, e As RoutedEventArgs) Handles btnprevious.Click
        index -= 1
        showData(index)
    End Sub

    Private Sub btnfirst_Click(sender As Object, e As RoutedEventArgs) Handles btnfirst.Click
        index = 0
        showData(index)
    End Sub

    Private Sub btnmovelast_Click(sender As Object, e As RoutedEventArgs) Handles btnmovelast.Click
        index = Table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub btndisplay_Click(sender As Object, e As RoutedEventArgs) Handles btndisplay.Click
        Dim report As New ReportViewer
        report.Show()
        Dim myrep As New Ripoti_Wards
        report.Title = "Ward Details"
        myrep.Load("..\\Ripoti_Wards.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub

    Private Sub wardDetails_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            index = table.Rows.Count()
            showData(index)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
