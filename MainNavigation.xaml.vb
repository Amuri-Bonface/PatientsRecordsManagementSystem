Public Class MainNavigation
    Private Sub MainNavigation_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        End
    End Sub

    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        Dim patients As New PatientsRegistration
        patients.lblstring.Content = lblconnstring.Content
        patients.Show()

    End Sub

    Private Sub buttonappointment_Click(sender As Object, e As RoutedEventArgs) Handles buttonappointment.Click
        Dim appointment As New Appointment
        appointment.lblstring.Content = lblconnstring.Content
        appointment.Show()
    End Sub

    Private Sub btnregdoctors_Click(sender As Object, e As RoutedEventArgs) Handles btnregdoctors.Click
        Dim daktari As New DoctorsRegister
        daktari.lblstring.Content = lblconnstring.Content
        daktari.Show()
    End Sub

    Private Sub btnregnurses_Click(sender As Object, e As RoutedEventArgs) Handles btnregnurses.Click
        Dim nurse As New NurseDetails
        nurse.lblstring.Content = lblconnstring.Content
        nurse.Show()
    End Sub

    Private Sub btnregwards_Click(sender As Object, e As RoutedEventArgs) Handles btnregwards.Click
        Dim regward As New wardDetails
        regward.lblstring.Content = lblconnstring.Content
        regward.Show()
    End Sub

    Private Sub btndiagnosis_Click(sender As Object, e As RoutedEventArgs) Handles btndiagnosis.Click
        Dim diagnosis As New PatientDiagnosis
        diagnosis.lblstring.Content = lblconnstring.Content
        diagnosis.Show()
    End Sub

    Private Sub btnwardadmission_Click(sender As Object, e As RoutedEventArgs) Handles btnwardadmission.Click
        Dim admitward As New WardAdmission
        admitward.lblstring.Content = lblconnstring.Content
        admitward.Show()
    End Sub

    Private Sub btnbilling_Click(sender As Object, e As RoutedEventArgs) Handles btnbilling.Click
        Dim bill As New BillingDetails
        bill.lblstring.Content = lblconnstring.Content
        bill.Show()
    End Sub

    Private Sub btnregusers_Click(sender As Object, e As RoutedEventArgs) Handles btnregusers.Click
        Dim add_user As New RegisterUser
        add_user.stringconn.Content = lblconnstring.Content
        add_user.Show()
    End Sub

    Private Sub btnstaff_Click(sender As Object, e As RoutedEventArgs) Handles btnstaff.Click
        Dim otherstaff As New Register_otherStaff
        otherstaff.lblstring.Content = lblconnstring.Content
        otherstaff.Show()
    End Sub

    Private Sub Button_Click_2(sender As Object, e As RoutedEventArgs)
        btnpatient.IsEnabled = False
        buttonappointment.IsEnabled = False
        btndiagnosis.IsEnabled = False
        btnwardadmission.IsEnabled = False
        btnbilling.IsEnabled = False
        btnregnurses.IsEnabled = False
        btnregwards.IsEnabled = False
        btnregdoctors.IsEnabled = False
        btnstaff.IsEnabled = False
        btnregusers.IsEnabled = False

        Dim logon As New Login
        logon.Show()
        logon.stringconn.Content = lblconnstring.Content
        Me.Hide()

    End Sub
End Class
