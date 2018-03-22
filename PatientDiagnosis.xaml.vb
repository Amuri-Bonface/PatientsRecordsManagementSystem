Imports System.Data
Imports System.Data.SqlClient
Public Class PatientDiagnosis
    Dim index As Integer = 0
    Dim table As New DataTable()

    Private Sub PatientDiagnosis_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader

            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Appointments"
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cboappointment.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()

            cmd2.Connection = cs
            cmd2.CommandText = "Select * from Nurse_Details"
            dr = cmd2.ExecuteReader

            With dr
                While .Read()
                    cbonurseID.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()


        Catch ex As Exception
        End Try

        index = table.Rows.Count()
        showData(index)
    End Sub


    Private Sub cboappointment_MouseLeave(sender As Object, e As MouseEventArgs) Handles cboappointment.MouseLeave
        Try
            Dim a As String
            a = cboappointment.Text

            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader

            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Appointments where Appointment_NO='" & Val(a) & "'"
            dr = cmd.ExecuteReader

            With dr
                .Read()
                txtpatientsID.Text = .GetValue(1)
                txtname.Text = .GetValue(2)
                txtdocID.Text = .GetValue(3)
                txtdocname.Text = .GetValue(4)

            End With
            dr.Close()
        Catch ex As Exception
        End Try
    End Sub


    Private Sub optyes_Click(sender As Object, e As RoutedEventArgs) Handles optyes.Click
        lbladmitted.Content = "Admitted"
        admissionCode.Content = 1
    End Sub

    Private Sub optno_Click(sender As Object, e As RoutedEventArgs) Handles optno.Click
        lbladmitted.Content = "Not Admitted"
        admissionCode.Content = 0
    End Sub


    Private Sub cbonurseID_MouseLeave(sender As Object, e As MouseEventArgs) Handles cbonurseID.MouseLeave
        Try
            Dim a As String
            a = cbonurseID.Text

            Dim connectionstring As String
            connectionstring = lblstring.Content

            Dim cs As New SqlConnection(connectionstring)
            Dim cmd, cmd2 As New SqlCommand
            Dim dr As SqlDataReader

            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Nurse_Details where Nurse_employment_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader

            With dr
                .Read()

                txtnurseName.Text = .GetValue(1)


            End With
            dr.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click

        txtdiagnosis.Text = ""
        cboappointment.Text = ""
        txtpatientsID.Text = ""
        txtname.Text = ""
        txtdocID.Text = ""
        txtdocname.Text = ""
        txtsymptoms.Text = ""
        txtcomments.Text = ""
        lbladmitted.Content = ""
        cbonurseID.Text = ""
        txtnurseName.Text = ""
        txtmedication.Text = ""

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

        cmd.CommandText = "Select * from Patients_Diagnosis"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(table)
        Dim last, newNum As Integer

        With table
            Try
                last = .Rows(position)(0)
                If last = 0 Then
                    txtdiagnosis.Text = 0
                Else
                    newNum = 1 + last
                    txtdiagnosis.Text = newNum
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
            cmd.CommandText = "insert into Patients_Diagnosis(Diagnosis_No,Appointment_NO,Patients_ID,Patients_name,Doctors_ID,Doctor_Name,Symptoms,Doctors_comments,Admission,Nurse_ID,Nurse_Name,Medication,AdmissionCode) Values
        ('" & txtdiagnosis.Text & "','" & cboappointment.Text & "','" & txtpatientsID.Text & "','" & txtname.Text & "','" & txtdocID.Text & "','" & txtdocname.Text & "',
        '" & txtsymptoms.Text & "','" & txtcomments.Text & "','" & lbladmitted.Content & "','" & cbonurseID.Text & "',
        '" & txtnurseName.Text & "','" & txtmedication.Text & "','" & Val(admissionCode.Content) & "')"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Patients Diagnosis updated Succesful", "Kapenguria Patients MIS", vbOK)
            cs.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        index = table.Rows.Count() - 1
        showData(index)
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Diagnosis. No to search")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Patients_Diagnosis where Diagnosis_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader
            Try
                With dr
                    .Read()
                    txtdiagnosis.Text = .GetValue(0)
                    cboappointment.Text = .GetValue(1)
                    txtpatientsID.Text = .GetValue(2)
                    txtname.Text = .GetValue(3)
                    txtdocID.Text = .GetValue(4)
                    txtdocname.Text = .GetValue(5)
                    txtsymptoms.Text = .GetValue(6)
                    txtcomments.Text = .GetValue(7)
                    lbladmitted.Content = .GetValue(8)
                    cbonurseID.Text = .GetValue(9)
                    txtnurseName.Text = .GetValue(10)
                    txtmedication.Text = .GetValue(11)

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
        Try
            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "update Patients_Diagnosis set Diagnosis_No='" & txtdiagnosis.Text & "',Appointment_NO='" & cboappointment.Text & "',Patients_ID='" & txtpatientsID.Text & "',Patients_name='" & txtname.Text & "',Doctors_ID='" & txtdocID.Text & "'
			,Doctor_Name='" & txtdocname.Text & "',Symptoms='" & txtsymptoms.Text & "',Doctors_comments='" & txtcomments.Text & "',Admission='" & lbladmitted.Content & "'
			,Nurse_ID='" & cbonurseID.Text & "',Nurse_Name='" & txtnurseName.Text & "',Medication='" & txtmedication.Text & "',AdmissionCode='" & admissionCode.Content & "'"

            cmd.ExecuteNonQuery()
            MessageBox.Show("Patients_Diagnosis Updated Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Diagnosis to Delete")

        Dim connectionstring As String
        connectionstring = lblstring.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Patients_Diagnosis where Diagnosis_No='" & Val(a) & "'"
            dr = cmd.ExecuteReader
            Try
                With dr
                    .Read()

                    txtdiagnosis.Text = .GetValue(0)
                    cboappointment.Text = .GetValue(1)
                    txtpatientsID.Text = .GetValue(2)
                    txtname.Text = .GetValue(3)
                    txtdocID.Text = .GetValue(4)
                    txtdocname.Text = .GetValue(5)
                    txtsymptoms.Text = .GetValue(6)
                    txtcomments.Text = .GetValue(7)
                    lbladmitted.Content = .GetValue(8)
                    cbonurseID.Text = .GetValue(9)
                    txtnurseName.Text = .GetValue(10)
                    txtmedication.Text = .GetValue(11)
                End With
                dr.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
                Case vbYes
                    ''Deleting current record
                    cmd.CommandText = "Delete from Patients_Diagnosis where Diagnosis_No='" & Val(a) & "'"
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

        cmd.CommandText = "Select * from Nurse_Details"
        Dim dr As New SqlDataAdapter(cmd)

        dr.Fill(Table)
        With table
            Try

                txtdiagnosis.Text = .Rows(position)(0).ToString
                cboappointment.Text = .Rows(position)(1).ToString
                txtpatientsID.Text = .Rows(position)(2).ToString
                txtname.Text = .Rows(position)(3).ToString
                txtdocID.Text = .Rows(position)(4).ToString
                txtdocname.Text = .Rows(position)(5).ToString
                txtsymptoms.Text = .Rows(position)(6).ToString
                txtcomments.Text = .Rows(position)(7).ToString
                lbladmitted.Content = .Rows(position)(8).ToString
                cbonurseID.Text = .Rows(position)(9).ToString
                txtnurseName.Text = .Rows(position)(10).ToString
                txtmedication.Text = .Rows(position)(11).ToString
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
        Dim myrep As New Ripoti_Diagnosis
        report.Title = "Patients Diagnosis"
        myrep.Load("..\\Ripoti_Diagnosis.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub

    Private Sub buttonAdd_Loaded(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Loaded
        Try
            index = table.Rows.Count()
            showData(index)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnDiagnonis_Click(sender As Object, e As RoutedEventArgs) Handles btnDiagnonis.Click
        Dim report As New ReportViewer
        report.Show()
        Dim myrep As New Ripoti_Patients_Diagnosis
        report.Title = "Patient Diagnosis"
        myrep.Load("..\\Ripoti_Patients_Diagnosis.rpt")
        report.CrystalReportsViewer.ViewerCore.ReportSource = myrep
    End Sub
End Class
