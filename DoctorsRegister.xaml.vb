Imports System.Data
Imports System.Data.SqlClient
Public Class DoctorsRegister
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        txtempno.Text = ""
        txtname.Text = ""
        txtdepartment.Text = ""
        cboavailable.Text = ""

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

        cmd.CommandText = "Select * from Doctors"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(3)
                If last = 0 Then
                    txtempno.Text = 0
                Else
                    newNum = (1 + last) + 1000
                    txtempno.Text = newNum
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

        Try
            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "insert into Doctors(Doctor_Employment_Number,Doctors_name,Department,Available) Values
('" & txtempno.Text & "','" & txtname.Text & "','" & txtdepartment.Text & "','" & cboavailable.Text & "')"

            cmd.ExecuteNonQuery()
            MessageBox.Show("Doctor Registered Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Employment. No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Doctors where Doctor_Employment_Number='" & Val(a) & "'"
            dr = cmd.ExecuteReader

            With dr
                .Read()

                txtname.Text = .GetValue(0)
                txtdepartment.Text = .GetValue(1)
                cboavailable.Text = .GetValue(2)
                txtempno.Text = .GetValue(3)
            End With
            dr.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
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

            cmd.CommandText = "update Doctors set Doctor_Employment_Number='" & txtempno.Text & "',Doctors_name='" & txtname.Text & "',Department='" & txtdepartment.Text & "',Available='" & cboavailable.Text & "'"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Doctor Updates Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Doctors ID to Delete")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Doctors where Doctor_Employment_Number='" & Val(a) & "'"
            dr = cmd.ExecuteReader
            Try
                With dr
                    .Read()

                    txtname.Text = .GetValue(0)
                    txtdepartment.Text = .GetValue(1)
                    cboavailable.Text = .GetValue(2)
                    txtempno.Text = .GetValue(3)
                End With
                dr.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
                Case vbYes
                    ''Deleting current record
                    cmd.CommandText = "Delete from Doctors where Doctor_Employment_Number='" & Val(a) & "'"
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("its been a pleasure Freeing up some space, Delete Succesfull")
                    cs.Close()
                Case vbNo
                    Exit Sub
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub showData(position As Integer)
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)

        Dim cmd, cmd2 As New SqlCommand
        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Doctors"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(Table)
        With Table
            Try
                txtname.Text = .Rows(position)(0).ToString
                txtdepartment.Text = .Rows(position)(1).ToString
                cboavailable.Text = .Rows(position)(2).ToString
                txtempno.Text = .Rows(position)(3).ToString

            Catch ex As Exception
                MessageBox.Show(ex.Message)
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
        Dim myrep As New Ripoti_Doctors
        report.Title = "Doctors Report"
        myrep.Load("..\\Ripoti_Doctors.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub

    Private Sub DoctorsRegister_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            index = table.Rows.Count()
            showData(index)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
