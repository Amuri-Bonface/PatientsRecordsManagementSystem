
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class ReportViewer

    Private Sub CrystalReportsViewer_Loaded(sender As Object, e As RoutedEventArgs) Handles CrystalReportsViewer.Loaded
        Dim cryRpt As New Ripoti_Appointment()
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table



        CrTables = cryRpt.Database.Tables
        For Each CrTable In CrTables
            crtableLogoninfo = CrTable.LogOnInfo
            crtableLogoninfo.ConnectionInfo = crConnectionInfo
            CrTable.ApplyLogOnInfo(crtableLogoninfo)
        Next
    End Sub

    Private Sub ReportViewer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

    End Sub
End Class
