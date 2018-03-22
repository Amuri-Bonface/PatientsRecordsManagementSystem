Imports System.Data
Imports System.Data.SqlClient
Public Class NurseDetails
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        txtempno.Text = ""
        txtname.Text = ""
        txtdepartment.Text = ""
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

        cmd.CommandText = "Select * from Nurse_Details"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtempno.Text = 0
                Else
                    newNum = (1 + last) + 100
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
        Try
            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "insert into Nurse_Details(Nurse_employment_No,Nurse_Name,Department) Values
('" & txtempno.Text & "','" & txtname.Text & "','" & txtdepartment.Text & "')"

            cmd.ExecuteNonQuery()
            MessageBox.Show("Nurse Registered Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception

        End Try
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

            cmd.CommandText = "Select * from Nurse_Details where Nurse_employment_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader

            With dr
                .Read()

                txtname.Text = .GetValue(1)
                txtdepartment.Text = .GetValue(2)
                txtempno.Text = .GetValue(0)
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

            cmd.CommandText = "update Nurse_Details set Nurse_employment_No='" & txtempno.Text & "',Nurse_Name='" & txtname.Text & "',Department='" & txtdepartment.Text & "'"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Nurse Updates Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Nurse ID to Delete")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        
        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Nurse_Details where Nurse_employment_No='" & Val(a) & "'"
        dr = cmd.ExecuteReader
        Try
            With dr
                .Read()

                txtname.Text = .GetValue(1)
                txtdepartment.Text = .GetValue(2)
                txtempno.Text = .GetValue(0)
            End With
            dr.Close()
        Catch ex As Exception

        End Try

        Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
            Case vbYes
                ''Deleting current record
                cmd.CommandText = "Delete from Nurse_Details where Nurse_employment_No='" & Val(a) & "'"
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

        cmd.CommandText = "Select * from Nurse_Details"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(Table)
        With Table
            Try
                txtname.Text = .Rows(position)(1).ToString
                txtdepartment.Text = .Rows(position)(2).ToString
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
        Dim myrep As New Ripoti_Nurses
        report.Title = "Nurse Details Report"
        myrep.Load("..\\Ripoti_Nurses.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub

    Private Sub NurseDetails_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            index = table.Rows.Count()
            showData(index)
        Catch ex As Exception

        End Try
    End Sub
End Class
