Imports System.Data.SqlClient
Imports System.IO

Class MainWindow
    Public dsource, catalog, user, pass, integrity As String
    Public connectionstring As String
    Private Sub CheckBox_Checked(sender As Object, e As RoutedEventArgs)

        If checkIntegrity.IsChecked Then
            lblintegrity.Content = "True"
            Try

                dsource = txtdatasource.Text
                catalog = txtcatalog.Text
                user = txtuser.Text
                pass = txtpassword.Password
                integrity = lblintegrity.Content
                connectionstring = "Data Source=" + dsource + ";" + "Initial Catalog=" + catalog + ";" + "Integrated Security=" + integrity
                txtstring.Text = connectionstring
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub


    Private Sub txtdatasource_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtdatasource.TextChanged
        Try
            Dim dsource, catalog, user, pass, integrity As String
            Dim connectionstring As String

            dsource = txtdatasource.Text
            catalog = txtcatalog.Text
            user = txtuser.Text
            pass = txtpassword.Password
            integrity = lblintegrity.Content
            connectionstring = "Data Source=" + dsource + ";" + "Initial Catalog=" + catalog + ";" + "User ID=" + user + ";" + "Password=" + pass
            txtstring.Text = connectionstring
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim connectionstring As String
        connectionstring = txtstring.Text

        Dim cs As New SqlConnection(connectionstring)
        Dim cmd, cmd2 As New SqlCommand
        Dim dr As SqlDataReader

        Try
            ''Deleting current configuration
            cmd.Connection = cs
            cs.Open()
            cmd.CommandText = "Delete FROM configuration"
            cmd.ExecuteNonQuery()
            cs.Close()


            cmd2.Connection = cs
            cs.Open()
            cmd.CommandText = "insert into configuration(DataSource,Catalog,User_Id,Password,Integrated_security,connection_string) values('" & txtdatasource.Text & "','" & txtcatalog.Text & "','" & txtuser.Text & "','" & txtpassword.Password & "','" & lblintegrity.Content & "','" & txtstring.Text & "')"
            ' cmd.CommandText = "update configuration set DataSource='" & txtdatasource.Text & "',Catalog='" & txtcatalog.Text & "',User_Id='" & txtuser.Text & "',Password='" & txtpassword.Password & "',Integrated_security='" & lblintegrity.Content & "',connection_string='" & txtstring.Text & "'"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Update Succesful", "Kapenguria Patients MIS", vbOK)

            cs.Close()


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        ''Outputting configuration to text file
        Dim strFile As String = "..\\Patients_config.txt"
        Dim sw As StreamWriter
        Dim fs As FileStream

        If (Not File.Exists(strFile)) Then
            Try
                fs = File.Create(strFile)
                sw = File.AppendText(strFile)
                sw.WriteLine(connectionstring)

            Catch ex As Exception

            End Try
        ElseIf (File.Exists(strFile)) Then

            sw = File.AppendText(strFile)
                sw.WriteLine(connectionstring)
            sw.Close()
        End If

    End Sub


    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        'retrieving configuration settings
        Try
            ' Open the file using a stream reader.
            Using sr As New StreamReader("..\\Patients_config.txt")
                Dim connecti As String
                ' Read the stream to a string and write the string to the console.
                connecti = sr.ReadToEnd()
                txtstring.Text = connecti
                sr.Close()

                ''connect to Database
                Dim cs As New SqlConnection(connecti)
                Dim cmd As New SqlCommand
                Dim dr As SqlDataReader
                Try
                    cmd.Connection = cs
                    cs.Open()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End Using

        Catch ex As Exception
            MessageBox.Show("The file could not be read:")

        End Try

    End Sub

    Private Sub buttonProceed_Click(sender As Object, e As RoutedEventArgs) Handles buttonProceed.Click
        Dim login As New Login

        login.Show()
        login.stringconn.Content = txtstring.Text
        Me.Hide()
    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        End
    End Sub
End Class
