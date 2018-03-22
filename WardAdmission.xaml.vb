Imports System.Data
Imports System.Data.SqlClient
Public Class WardAdmission
    Dim index As Integer = 0
    Dim table As New DataTable()
    Private Sub cbodiagnosis_MouseLeave(sender As Object, e As MouseEventArgs) Handles cbodiagnosis.MouseLeave
        Try
            Dim a As Integer
            a = cbodiagnosis.Text

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
                    txtpatientsID.Text = .GetValue(2)
                    txtname.Text = .GetValue(3)
                    txtdocID.Text = .GetValue(4)
                    txtdocname.Text = .GetValue(5)
                    txtnurseID.Text = .GetValue(9)
                    txtnurseName.Text = .GetValue(10)
                End With
                dr.Close()
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub

    Private Sub WardAdmission_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader


        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Patients_Diagnosis"
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cbodiagnosis.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()
        Catch ex As Exception
        End Try


        cmd2.Connection = cs
            cmd2.CommandText = "Select * from Wards"
            dr = cmd2.ExecuteReader
        Try
            With dr
                While .Read()
                    cbowardno.Items.Add(.GetValue(0))
                End While
            End With
        Catch ex As Exception
        End Try

        dr.Close()

        index = table.Rows.Count()
        showData(index)
    End Sub


    Private Sub cbowardno_MouseLeave(sender As Object, e As MouseEventArgs) Handles cbowardno.MouseLeave
        Try
            Dim a As Integer
            a = cbowardno.Text

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
                    txtwardname.Text = .GetValue(2)

                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            dr.Close()

            Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        txtadmissioNo.Text = ""
        cbodiagnosis.Text = ""
        txtpatientsID.Text = ""
        txtname.Text = ""
        txtdocID.Text = ""
        txtdocname.Text = ""
        txtnurseID.Text = ""
        txtnurseName.Text = ""
        cbowardno.Text = ""
        txtwardname.Text = ""
        txtbedno.Text = ""
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

        cmd.CommandText = "Select * from Ward_Admission"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtadmissioNo.Text = 0
                Else
                    newNum = 1 + last
                    txtadmissioNo.Text = newNum
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
            cmd.CommandText = "insert into Ward_Admission(Ward_admission_No,Diagnosis_No,Patients_ID,Patients_name,Doc_ID,Doc_Name,Nurse_ID,Nurse_Name,Ward_No,Ward_Name,Bed_No,admission_date,Hospital_Bill) Values
        ('" & txtadmissioNo.Text & "','" & cbodiagnosis.Text & "','" & txtpatientsID.Text & "','" & txtname.Text & "','" & txtdocID.Text & "','" & txtdocname.Text & "',
        '" & txtnurseID.Text & "','" & txtnurseName.Text & "','" & cbowardno.Text & "','" & txtwardname.Text & "',
        '" & txtbedno.Text & "','" & dtpicker.Text & "','" & txthosBill.Text & "')"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Succesfully Admitted to ward", "Kapenguria Patients MIS", vbOK)
            cs.Close()

        Catch ex As Exception

        End Try
        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Ward_admission_No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Ward_Admission where Ward_admission_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader
            Try
                With dr
                    .Read()
                    txtadmissioNo.Text = .GetValue(0)
                    cbodiagnosis.Text = .GetValue(1)
                    txtpatientsID.Text = .GetValue(2)
                    txtname.Text = .GetValue(3)
                    txtdocID.Text = .GetValue(4)
                    txtdocname.Text = .GetValue(5)
                    txtnurseID.Text = .GetValue(6)
                    txtnurseName.Text = .GetValue(7)
                    cbowardno.Text = .GetValue(8)
                    txtwardname.Text = .GetValue(9)
                    txtbedno.Text = .GetValue(10)
                    dtpicker.Text = .GetValue(11)
                    txthosBill.Text = .GetValue(12)

                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

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


        cmd.Connection = cs
        cs.Open()
        Try
            cmd.CommandText = "update Ward_Admission set Ward_admission_No='" & txtadmissioNo.Text & "',Diagnosis_No='" & cbodiagnosis.Text & "',Patients_ID='" & txtpatientsID.Text & "',Patients_name='" & txtname.Text & "',Doc_ID='" & txtdocID.Text & "'
			,Doc_Name='" & txtdocname.Text & "',Nurse_ID='" & txtnurseID.Text & "',Nurse_Name='" & txtnurseName.Text & "',Ward_No='" & cbowardno.Text & "'
			,Ward_Name='" & txtwardname.Text & "',Bed_No='" & txtbedno.Text & "',admission_date='" & dtpicker.Text & "',Hospital_Bill='" & txthosBill.Text & "'"
            cmd.ExecuteNonQuery()
            MessageBox.Show(" Succesful Updated Ward Admission", "Kapenguria Patients MIS", vbOK)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        cs.Close()
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Ward Admission No to Delete")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()
        Try
            cmd.CommandText = "Select * from Ward_Admission where Ward_admission_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader
            Try
                With dr
                    .Read()

                    txtadmissioNo.Text = .GetValue(0)
                    cbodiagnosis.Text = .GetValue(1)
                    txtpatientsID.Text = .GetValue(2)
                    txtname.Text = .GetValue(3)
                    txtdocID.Text = .GetValue(4)
                    txtdocname.Text = .GetValue(5)
                    txtnurseID.Text = .GetValue(6)
                    txtnurseName.Text = .GetValue(7)
                    cbowardno.Text = .GetValue(8)
                    txtwardname.Text = .GetValue(9)
                    txtbedno.Text = .GetValue(10)
                    dtpicker.Text = .GetValue(11)
                    txthosBill.Text = .GetValue(12)
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            dr.Close()

        Catch ex As Exception
        End Try

        Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
            Case vbYes
                ''Deleting current record
                cmd.CommandText = "Delete from Ward_Admission where Ward_admission_No='" & Val(a) & "'"
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

        cmd.CommandText = "Select * from Ward_Admission"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        With table
            Try
                txtadmissioNo.Text = .Rows(position)(0).ToString
                cbodiagnosis.Text = .Rows(position)(1).ToString
                txtpatientsID.Text = .Rows(position)(2).ToString
                txtname.Text = .Rows(position)(3).ToString
                txtdocID.Text = .Rows(position)(4).ToString
                txtdocname.Text = .Rows(position)(5).ToString
                txtnurseID.Text = .Rows(position)(6).ToString
                txtnurseName.Text = .Rows(position)(7).ToString
                cbowardno.Text = .Rows(position)(8).ToString
                txtwardname.Text = .Rows(position)(9).ToString
                txtbedno.Text = .Rows(position)(10).ToString
                dtpicker.Text = .Rows(position)(11).ToString
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
        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub btndisplay_Click(sender As Object, e As RoutedEventArgs) Handles btndisplay.Click
        Dim report As New ReportViewer
        report.Show()
        Dim myrep As New Ripoti_WardAdmission
        report.Title = "Ward Admission Report"
        myrep.Load("..\\Ripoti_WardAdmission.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub


End Class
