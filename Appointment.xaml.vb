Imports System.Data
Imports System.Data.SqlClient
Public Class Appointment
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub Appointment_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Patients_registration"
        dr = cmd.ExecuteReader

        With dr
            While .Read()
                cbopatientsID.Items.Add(.GetValue(5))
            End While
        End With
        dr.Close()

        'retrive doctors
        cmd.Connection = cs
        cmd.CommandText = "Select * from Doctors"
        dr = cmd.ExecuteReader
        With dr
            While .Read()
                txtdocID.Items.Add(.GetValue(3))
            End While
        End With
        dr.Close()

        'make form display some data
        index = table.Rows.Count()
        showData(index)
    End Sub


    Private Sub cbopatientsID_MouseLeave(sender As Object, e As MouseEventArgs) Handles cbopatientsID.MouseLeave
        Try
            Dim a As String
            a = cbopatientsID.Text

            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader

            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from patients_registration where Id_no='" & Val(a) & "'"
            dr = cmd.ExecuteReader

            With dr
                .Read()

                txtname.Text = .GetValue(0)

            End With
            dr.Close()
        Catch ex As Exception
        End Try
    End Sub



    Private Sub txtdocID_MouseLeave(sender As Object, e As MouseEventArgs) Handles txtdocID.MouseLeave
        Try

            Dim a As String

            a = txtdocID.Text

            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader

            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Doctors where Doctor_Employment_Number='" & Val(a) & "'"
            dr = cmd.ExecuteReader

            With dr
                .Read()

                txtdocname.Text = .GetValue(0)

            End With
            dr.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        txtappointmentno.Text = ""
        cbopatientsID.Text = ""
        txtname.Text = ""
        txtdocID.Text = ""
        txtdocname.Text = ""
        dtpicker.Text = ""
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

        cmd.CommandText = "Select * from Appointments"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtappointmentno.Text = 0
                Else
                    newNum = 1 + last
                    txtappointmentno.Text = newNum
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
            cmd.CommandText = "insert into Appointments(Appointment_No,Patients_ID,Patients_Name,Doctor_ID,Doctor_Name,Date) Values
        ('" & txtappointmentno.Text & "','" & cbopatientsID.Text & "','" & txtname.Text & "','" & txtdocID.Text & "','" & txtdocname.Text & "','" & dtpicker.Text & "')"

            cmd.ExecuteNonQuery()
            MessageBox.Show("Appointment Booked Succesfully", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Appointment No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Appointments where Appointment_No='" & Val(a) & "'"
        dr = cmd.ExecuteReader

        With dr
            .Read()

            txtappointmentno.Text = .GetValue(0)
            cbopatientsID.Text = .GetValue(1)
            txtname.Text = .GetValue(2)
            txtdocID.Text = .GetValue(3)
            txtdocname.Text = .GetValue(4)
            dtpicker.Text = .GetValue(5)

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

        cmd.CommandText = "update Appointments set Appointment_No='" & txtappointmentno.Text & "',Patients_ID='" & cbopatientsID.Text & "',Patients_Name='" & txtname.Text & "',
Doctor_Name='" & txtdocname.Text & "',Date='" & dtpicker.Text & "'"
        cmd.ExecuteNonQuery()
        MessageBox.Show("Appointment Updated Succesful", "Kapenguria Patients MIS", vbOK)

        cs.Close()
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Appointment No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Appointments where Appointment_No='" & Val(a) & "'"
        dr = cmd.ExecuteReader
        Try
            With dr
                .Read()

                txtappointmentno.Text = .GetValue(0)
                cbopatientsID.Text = .GetValue(1)
                txtname.Text = .GetValue(2)
                txtdocID.Text = .GetValue(3)
                txtdocname.Text = .GetValue(4)
                dtpicker.Text = .GetValue(5)

            End With
            dr.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Select Case MessageBox.Show("Please confirm you want to delete this record", "Freeing Up Some Space", vbYesNo)
            Case vbYes
                ''Deleting current record
                cmd.CommandText = "Delete from Appointments where Appointment_No='" & Val(a) & "'"
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

        cmd.CommandText = "Select * from Appointments"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        With table
            Try
                txtappointmentno.Text = .Rows(position)(0).ToString
                cbopatientsID.Text = .Rows(position)(1).ToString
                txtname.Text = .Rows(position)(2).ToString
                txtdocID.Text = .Rows(position)(3).ToString
                txtdocname.Text = .Rows(position)(4).ToString
                dtpicker.Text = .Rows(position)(5).ToString
            Catch ex As Exception
                MessageBox.Show(ex.Message)
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
        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub btndisplay_Click(sender As Object, e As RoutedEventArgs) Handles btndisplay.Click
        Dim report As New ReportViewer
        report.Show()
        Dim myrep As New Ripoti_Appointment

        myrep.Load("..\\Ripoti_Appointment.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub
End Class
