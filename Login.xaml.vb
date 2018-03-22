Imports System.Data.SqlClient
Public Class Login
    Private Sub cbouserlevel_MouseLeave(sender As Object, e As MouseEventArgs) Handles cbouserlevel.MouseLeave
        cboempno.Items.Clear()
        Dim connectionstring As String
        connectionstring = stringconn.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader
        Dim levels As String
        levels = cbouserlevel.Text

        If levels = "Management" Then
            Try
                cmd.Connection = cs
                cs.Open()

                cmd.CommandText = "Select * from Login_Details where CONVERT(VARCHAR, User_Type)='" & levels & "'"
                dr = cmd.ExecuteReader

                With dr
                    While .Read()
                        cboempno.Items.Add(.GetValue(2))
                    End While
                End With
                dr.Close()
            Catch ex As Exception

            End Try
        ElseIf levels = "Receptionist" Then
            Try
                cmd.Connection = cs
                cs.Open()

                cmd.CommandText = "Select * from Login_Details where CONVERT(VARCHAR, User_Type)='" & levels & "'"
                dr = cmd.ExecuteReader

                With dr
                    While .Read()
                        cboempno.Items.Add(.GetValue(2))
                    End While
                End With
                dr.Close()

            Catch ex As Exception

            End Try
        ElseIf levels = "Billing Officer" Then
            Try
                cmd.Connection = cs
                cs.Open()

                cmd.CommandText = "Select * from Login_Details where CONVERT(VARCHAR, User_Type)='" & levels & "'"
                dr = cmd.ExecuteReader

                With dr
                    While .Read()
                        cboempno.Items.Add(.GetValue(2))
                    End While
                End With
                dr.Close()

            Catch ex As Exception

            End Try
        ElseIf levels = "Nurse" Then
            Try
                cmd.Connection = cs
                cs.Open()

                cmd.CommandText = "Select * from Login_Details where CONVERT(VARCHAR, User_Type)='" & levels & "'"
                dr = cmd.ExecuteReader

                With dr
                    While .Read()
                        cboempno.Text = ""
                        cboempno.Items.Add(.GetValue(2))
                    End While
                End With
                dr.Close()

            Catch ex As Exception

            End Try
        ElseIf levels = "Doctor" Then
            Try
                cmd.Connection = cs
                cs.Open()

                cmd.CommandText = "Select * from Login_Details where CONVERT(VARCHAR, User_Type)='" & levels & "'"
                dr = cmd.ExecuteReader

                With dr
                    While .Read()

                        cboempno.Items.Add(.GetValue(2))
                    End While
                End With
                dr.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub btnlogin_Click(sender As Object, e As RoutedEventArgs) Handles btnlogin.Click
        Dim connectionstring As String
        connectionstring = stringconn.Content

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd As New SqlCommand
        Dim dr As SqlDataReader
        Dim userType, empNo, pass1, pass2 As String
        Try
            userType = cbouserlevel.Text
            empNo = cboempno.Text
            pass2 = txtpassword.Password

            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "Select * from Login_Details where CONVERT(VARCHAR, User_Type)='" & userType & "' AND CONVERT(VARCHAR, EmploymentNo)='" & empNo & "'"
            dr = cmd.ExecuteReader

            While dr.Read
                Try
                    pass1 = dr.GetValue(1)
                    If pass1 <> pass2 Then
                        MessageBox.Show("Wrong Password,Try again", "Login Form", vbOK)
                    Else
                        MessageBox.Show("Login Succesful", "Login Form", vbOK)
                        Dim main As New MainNavigation
                        main.Show()
                        main.lblconnstring.Content = stringconn.Content
                        If userType = "Management" Then
                            main.btnregnurses.IsEnabled = True
                            main.btnregdoctors.IsEnabled = True
                            main.btnregusers.IsEnabled = True
                            main.btnregwards.IsEnabled = True
                            main.btnstaff.IsEnabled = True
                        ElseIf userType = "Doctor" Then
                            main.btndiagnosis.IsEnabled = True
                        ElseIf userType = "Nurse" Then
                            main.btnwardadmission.IsEnabled = True
                        ElseIf userType = "Receptionist" Then
                            main.btnpatient.IsEnabled = True
                            main.buttonappointment.IsEnabled = True
                        ElseIf userType = "Billing Officer" Then
                            main.btnbilling.IsEnabled = True
                        End If
                        Me.Hide()
                    End If
                Catch ex As Exception

                End Try
            End While

        Catch ex As Exception

        End Try
    End Sub
End Class
