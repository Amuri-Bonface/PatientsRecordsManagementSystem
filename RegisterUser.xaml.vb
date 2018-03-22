Imports System.Data.SqlClient
Public Class RegisterUser
    Private Sub cbouserlevel_MouseLeave(sender As Object, e As MouseEventArgs) Handles cbouserlevel.MouseLeave

        Dim connectionstring As String
        connectionstring = stringconn.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Dim levels As String
        levels = cbouserlevel.Text

        If levels = "Management" Then
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from other_staff where CONVERT(VARCHAR, Duty)='" & levels & "'"
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cboempno.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()

        ElseIf levels = "Receptionist" Then
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from other_staff where CONVERT(VARCHAR, Duty)='" & levels & "'"
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cboempno.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()


        ElseIf levels = "Billing Officer" Then
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from other_staff where CONVERT(VARCHAR, Duty)='" & levels & "'"
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cboempno.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()


        ElseIf levels = "Nurse" Then
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Nurse_Details "
            dr = cmd.ExecuteReader

            With dr
                While .Read()
                    cboempno.Text = ""
                    cboempno.Items.Add(.GetValue(0))
                End While
            End With
            dr.Close()


        ElseIf levels = "Doctor" Then
            cmd.Connection = cs
            cs.Open()

            cmd.CommandText = "Select * from Doctors "
            dr = cmd.ExecuteReader

            With dr
                While .Read()

                    cboempno.Items.Add(.GetValue(3))
                End While
            End With
            dr.Close()
        End If
    End Sub

    Private Sub buttonAdd_Click(sender As Object, e As RoutedEventArgs) Handles buttonAdd.Click
        cbouserlevel.Text = ""
        cboempno.Text = ""
    End Sub

    Private Sub buttonSave_Click(sender As Object, e As RoutedEventArgs) Handles buttonSave.Click
        Dim connectionstring As String
        'getting the connection string from the Main Navigation form
        Dim mainConn As New MainNavigation
        connectionstring = stringconn.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Try
            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "insert into Login_Details(EmploymentNo,User_Type,Password) Values
('" & Val(cboempno.Text) & "','" & cbouserlevel.Text & "','" & txtpassword.Password & "')"

            cmd.ExecuteNonQuery()
            MessageBox.Show("User Registered Succesful ", "Kapenguria Patients MIS", vbOK)

            cs.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub buttonSearch_Click(sender As Object, e As RoutedEventArgs) Handles buttonSearch.Click
        Dim a As Integer
        a = InputBox("Enter Employment. No to search")

        Dim connectionstring As String
        connectionstring = stringconn.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Login_Details where EmploymentNo='" & Val(a) & "'"
        dr = cmd.ExecuteReader

        With dr
            .Read()

            cbouserlevel.Text = .GetValue(0)
            cboempno.Text = .GetValue(2)

        End With
        dr.Close()

    End Sub

    Private Sub buttonUpdate_Click(sender As Object, e As RoutedEventArgs) Handles buttonUpdate.Click
        Dim connectionstring As String
        connectionstring = stringconn.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "update Login_Details set EmploymentNo='" & cboempno.Text & "',User_Type='" & cbouserlevel.Text & "',Password='" & txtpassword.Password & "'"
        cmd.ExecuteNonQuery()
        MessageBox.Show("Update Succesful", "Kapenguria Patients MIS", vbOK)

        cs.Close()
    End Sub

    Private Sub buttonDelete_Click(sender As Object, e As RoutedEventArgs) Handles buttonDelete.Click
        Dim a As Integer
        a = InputBox("Enter Staff Employment No..to Delete")

        Dim connectionstring As String
        connectionstring = stringconn.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        cmd.Connection = cs
        cs.Open()

        cmd.CommandText = "Select * from Login_Details where EmploymentNo='" & Val(a) & "'"
        dr = cmd.ExecuteReader
        Try
            With dr
                .Read()

                cbouserlevel.Text = .GetValue(0)
                cboempno.Text = .GetValue(2)
            End With
            dr.Close()
        Catch ex As Exception

        End Try

        Select Case MessageBox.Show("Please confirm you delete this record", "Freeing Up Some Space", vbYesNo)
            Case vbYes
                ''Deleting current record
                cmd.CommandText = "Delete from Login_Details where EmploymentNo='" & Val(a) & "'"
                cmd.ExecuteNonQuery()
                MessageBox.Show("its been a pleasure Freeing up some space, Delete Succesfull")
                cs.Close()
            Case vbNo
                Exit Sub
        End Select

    End Sub
End Class
