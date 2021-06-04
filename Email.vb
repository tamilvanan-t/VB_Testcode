Imports System.Net.Mail
Public Class SendEmail
    Public Sub sendEmail(ByRef prop As Properties, ByVal message As String, ByVal attachmentFile As String)
        Dim from As String = prop.GetProperty("FROM")
        Dim toMail As String = prop.GetProperty("TO")
        Dim ccMail As String = prop.GetProperty("CC")
        Dim pasword As String = prop.GetProperty("PASSWORD")
        Dim smtpHost As String = prop.GetProperty("SMTP_HOST")
        Dim smtpPort As String = prop.GetProperty("SMTP_PORT")
        Dim subject As String = prop.GetProperty("SUBJECT")


        Dim toMailArray() As String
        toMailArray = toMail.Split(";")

        Dim ccMailArray() As String
        ccMailArray = ccMail.Split(";")

        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(from, pasword)
            Smtp_Server.Port = smtpPort
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = smtpHost

            e_mail = New MailMessage()
            e_mail.From = New MailAddress(from)
            e_mail.Subject = subject
            e_mail.IsBodyHtml = False
            e_mail.Body = message

            For Each toMailId As String In toMailArray
                e_mail.To.Add(toMailId)
            Next

            For Each ccMailId As String In ccMailArray
                If Not String.IsNullOrWhiteSpace(ccMailId) Then
                    e_mail.CC.Add(ccMailId)
                End If
            Next

            'Send Attachment
            Dim dirs As String() = attachmentFile.Split("\")
            Dim fileName As String = dirs(dirs.Length - 1)

            Dim data As Net.Mail.Attachment = New Net.Mail.Attachment(attachmentFile)
            data.Name = fileName
            e_mail.Attachments.Add(data)

            Smtp_Server.Send(e_mail)
        Catch error_t As Exception
            Console.WriteLine(error_t.ToString)
        End Try
    End Sub
End Class