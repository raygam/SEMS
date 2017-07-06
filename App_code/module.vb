Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Text
Imports System.Net.Mail

Public Class universalmodule

    Private Shared Function FileThere(ByRef f As String) As Boolean
        FileThere = File.Exists(f)
    End Function

    Public Shared Function SendEmail(ByRef EmailTo As String, ByRef EmailFrom As String, ByRef FromName As String, ByVal Subject As String, ByVal MessageHTML As String, ByVal SMTPServer As String, ByVal Attachment As String) As Boolean

        'On Error GoTo errhand
        Dim Emails() As String
        Dim i As Integer
        Emails = Split(EmailTo, ";")
        Dim AttachmentFile As Attachment
        Dim mail As New MailMessage
        mail.From = New MailAddress(EmailFrom, FromName)

        For i = 0 To UBound(Emails)
            mail.To.Add(Emails(i))
        Next

        mail.Bcc.Add("hdunn@smithers.com")
        mail.Subject = Subject
        mail.Body = MessageHTML
        mail.IsBodyHtml = True

        'If Attachment > "" Then
        Dim stream As FileStream = File.OpenRead(Attachment)
            Dim FriendlyFileName As String = Attachment.Substring(InStrRev(Attachment, "\"))
            AttachmentFile = New Attachment(stream, FriendlyFileName)
            mail.Attachments.Add(AttachmentFile)


        'End If

        Dim smtp As New SmtpClient(SMTPServer)
        smtp.Send(mail)
        SendEmail = True
        'If Attachment > "" Then
        'AttachmentFile.Dispose()
        AttachmentFile = Nothing
        Stream.Dispose()
        Stream = Nothing
        'End If
        mail.Dispose()
        'smtp.Dispose()
        mail = Nothing
        smtp = Nothing

        Exit Function
errhand:
        mail = Nothing
        smtp = Nothing
        SendEmail = False
    End Function

    Public Shared Sub AddTextLine(ByVal fs As FileStream, ByVal value As String)
        value = value & vbCrLf
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(value)
        fs.Write(info, 0, info.Length)
        fs = Nothing
    End Sub
    Public Shared Function FTPFile(ByRef FTPLocation As String, ByRef Username As String, ByRef Password As String, ByRef SourceFile As String) As Boolean
        On Error GoTo errhand
        Dim clsRequest As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create(FTPLocation), System.Net.FtpWebRequest)
        clsRequest.Credentials = New System.Net.NetworkCredential(Username, Password)
        clsRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile

        ' read in file...
        Dim bFile() As Byte = System.IO.File.ReadAllBytes(SourceFile)

        ' upload file...
        Dim clsStream As System.IO.Stream = clsRequest.GetRequestStream()
        clsStream.Write(bFile, 0, bFile.Length)
        clsStream.Close()
        clsStream.Dispose()
        clsStream = Nothing
        clsRequest = Nothing
        bFile = Nothing
        FTPFile = True
        Exit Function
errhand:
        clsStream = Nothing
        clsRequest = Nothing
        bFile = Nothing
        FTPFile = False
        'MessageBox.Show(Err.Description)
    End Function

    Public Shared Function GetDateFilename() As String

        Dim y As String
        y = Year(Date.Now)
        y = y.Substring(y.Length - 2, 2)

        Dim m As String
        m = "0" & Month(Date.Now)
        m = m.Substring(m.Length - 2, 2)

        Dim d As String
        d = "0" & Microsoft.VisualBasic.DateAndTime.Day(Date.Now)
        d = d.Substring(d.Length - 2, 2)

        GetDateFilename = y & m & d

    End Function
End Class
