Imports Arbitext.RegistryHelpers

Public Class EmailHelpers
    Public Shared Sub sendSilentNotification(emailBodyMessage As String, emailSubject As String)
        If ThisAddIn.EmailsOK Then
            Dim iMsg As Object
            Dim iConf As Object
            Dim strBody As String
            Dim Flds As Object
            iMsg = CreateObject("CDO.Message")
            iConf = CreateObject("CDO.Configuration")
            'iConf.Load -1    ' CDO Source Defaults
            Flds = iConf.Fields
            With Flds
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = ThisAddIn.EmailAddress
                .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ThisAddIn.EmailPassword
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
                .Update
            End With
            With iMsg
                .Configuration = iConf
                .To = ThisAddIn.EmailAddress
                .CC = ""
                .BCC = ""
                .from = ThisAddIn.EmailAddress
                .Subject = emailSubject
                .HTMLBody = emailBodyMessage
                On Error Resume Next
                .send
                On Error GoTo 0
            End With
        End If
    End Sub


End Class
