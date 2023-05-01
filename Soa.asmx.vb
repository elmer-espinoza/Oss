Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Web
Imports System.IO


' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class Soa
    Inherits System.Web.Services.WebService

    'Dim ConStr As String = "Provider=SQLOLEDB; Data Source=174.36.22.84\SQL2005,6688;Initial Catalog=Oss;Persist Security Info=True;User ID=elmer.espinoza;Password=P@ssw0rd"
    'Dim ConStr As String = "Provider=SQLOLEDB; Data Source=sv31.dbsqlserver.com,6688;Initial Catalog=Oss;Persist Security Info=True;User ID=elmer.espinoza;Password=P@ssw0rd"
    Dim ConStr As String = "Provider=SQLOLEDB; Data Source=sv32.dbsqlserver.com,8888;Initial Catalog=Oss;Persist Security Info=True;User ID=elmer.espinoza;Password=P@ssw0rd"

    <WebMethod()>
    Public Function HelloWorld() As String
        Return "Hola a todos"
    End Function


    <WebMethod()>
    Public Function Hola() As String
        Return "Hola Mundo la fecha hora es " + Now().ToString
    End Function


    <WebMethod()>
    Public Function Add(ByVal Numero1 As Integer, ByVal Numero2 As Integer) As Integer
        Return Numero1 + Numero2
    End Function


    <WebMethod()>
    Public Function Subtract(ByVal Numero1 As Integer, ByVal Numero2 As Integer) As Integer
        Return Numero1 - Numero2
    End Function


    <WebMethod()>
    Public Function CelciusToFahrenheit(Celcius As Double) As String
        Return Strings.Format(1.8 * Celcius + 32.0, "#0.0") + " °F"
    End Function


    <WebMethod()>
    Public Function Query(SQL As String) As DataSet
        Dim dataSet As New DataSet()
        Dim oleDbDataAdapter As New OleDbDataAdapter(SQL, New OleDbConnection(Me.ConStr))
        oleDbDataAdapter.Fill(dataSet)
        Return dataSet
    End Function


    <WebMethod()>
    Public Function GetData(ByVal Scope As String) As Data.DataSet
        Dim ds As New Data.DataSet
        'Dim da As New Data.OleDb.OleDbDataAdapter("select id, fecha, nombre, empresa, email, telefono, comentario, ip, browser from contactos where nombre+'-'+empresa+'-'+substring(comentario,1,250) like '%" + Scope + "%'", New Data.OleDb.OleDbConnection(ConStr))
        Dim da As New Data.OleDb.OleDbDataAdapter("SELECT distinct [equipo], [bandera], [grupo] FROM [oss].[dbo].[equipos]  where [equipo]+[bandera]+[grupo] like '%" + Scope + "%' order by equipo ", New Data.OleDb.OleDbConnection(ConStr))
        da.Fill(ds)
        Return ds
    End Function

    <WebMethod()>
    Public Function UpdateData(ByVal IP As String) As Boolean
        Try
            Dim cmd As New Data.OleDb.OleDbCommand
            Dim Conn As New Data.OleDb.OleDbConnection
            Conn.ConnectionString = ConStr
            Conn.Open()
            cmd.CommandText = "UPDATE CONTACTOS SET IP = '" + IP + "' "
            cmd.Connection = Conn
            cmd.ExecuteNonQuery()
            Conn.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    <WebMethod()>
    Public Function SendMail(ByVal Destination As String, ByVal Body As String, ByVal Encrypt As Boolean) As String 'Boolean
        Dim Status As String = "Mail was sent." 'Boolean = True
        Try
            Dim oMessage As New System.Net.Mail.MailMessage
            oMessage.Subject = "Automatic Mail"
            If Not Encrypt Then
                oMessage.Body = "<font face='Arial' color='#000000'><i>" + Body + "</i><br><br><b>Sent: " + Format(Now(), "yyyy-MM-dd HH:mm:ss.fff") + "</b><br><p><a href='http://www.ossperu.com'>http://www.ossperu.com</a></p></font>"
            Else
                oMessage.Body = EncodeBase64(Body + Format(Now(), "yyyyMMdd HH:mm:ss.fff"))
            End If

            oMessage.IsBodyHtml = True
            oMessage.Priority = System.Net.Mail.MailPriority.Normal
            Dim oemailClient As New System.Net.Mail.SmtpClient
            Dim basicAuthenticationInfo As New System.Net.NetworkCredential
            If False Then
                'Dim basicAuthenticationInfo As New System.Net.NetworkCredential("extranetperu@brinkster.net", "poison")
                oMessage.From = New System.Net.Mail.MailAddress("extranetperu@brinkster.net", "Web Master")
                basicAuthenticationInfo.UserName = "extranetperu@brinkster.net"
                basicAuthenticationInfo.Password = "poison"
                oemailClient.Host = "sendmail.brinkster.com"
            ElseIf True Then
                'Dim basicAuthenticationInfo As New System.Net.NetworkCredential("web.master@ossperu.com", "P@ssw0rd")
                oMessage.From = New System.Net.Mail.MailAddress("web.master@ossperu.com", "Web Master")
                basicAuthenticationInfo.UserName = "web.master@ossperu.com"
                'basicAuthenticationInfo.Password = "P@ssw0rd"
                'basicAuthenticationInfo.Password = "P@ssw0rd=="
                'basicAuthenticationInfo.Password = "Password123$"
                basicAuthenticationInfo.Password = "J0e_S@tr1ani"
                oemailClient.Host = "mail.ossperu.com"
            Else
                oMessage.From = New System.Net.Mail.MailAddress("web.master@xerox.com", "Web Master")
                oemailClient.Host = "forwarder.mail.xerox.com"
            End If

            oMessage.To.Add(New System.Net.Mail.MailAddress(Destination))
            oMessage.Bcc.Add(New System.Net.Mail.MailAddress("extranetperu@gmail.com", "Extranet"))

            oemailClient.EnableSsl = False
            oemailClient.Port = 25
            oemailClient.UseDefaultCredentials = False
            oemailClient.Credentials = basicAuthenticationInfo
            oemailClient.Send(oMessage)

        Catch ex As Exception
            Status = ex.Message 'False
        End Try
        Return Status
    End Function


    <WebMethod()>
    Public Function SendMailEx(ByVal Destination As String, ByVal Body As String, ByVal Encrypt As Boolean, ByVal FileName As String, ByVal FileBase64 As String) As String 'Boolean

        Dim Status As String = "Mail was Sent." 'Boolean = True

        Try
            Dim oMessage As New System.Net.Mail.MailMessage
            oMessage.Subject = "Automatic Mail "
            If Not Encrypt Then
                oMessage.Body = "<font face='Arial' color='#000000'><i>" + Body + "</i><br><br><b>Sent: " + Format(Now(), "yyyy-MM-dd HH:mm:ss.fff") + "</b><br><p><a href='http://www.ossperu.com'>http://www.ossperu.com</a></p></font>"
            Else
                oMessage.Body = EncodeBase64(Body + Format(Now(), "yyyyMMdd HH:mm:ss.fff"))
            End If

            oMessage.IsBodyHtml = True
            oMessage.Priority = System.Net.Mail.MailPriority.Normal
            Dim oemailClient As New System.Net.Mail.SmtpClient
            Dim basicAuthenticationInfo As New System.Net.NetworkCredential

            oemailClient.EnableSsl = False
            oemailClient.Port = 25
            oemailClient.UseDefaultCredentials = False


            If False Then
                'Dim basicAuthenticationInfo As New System.Net.NetworkCredential("extranetperu@brinkster.net", "poison")
                oMessage.From = New System.Net.Mail.MailAddress("extranetperu@brinkster.net", "Web Master")
                basicAuthenticationInfo.UserName = "extranetperu@brinkster.net"
                basicAuthenticationInfo.Password = "poison"
                oemailClient.Credentials = basicAuthenticationInfo
                oemailClient.Host = "sendmail.brinkster.com"
            ElseIf True Then ' No esta habilitado puerto SMTP en ossperu.com
                'Dim basicAuthenticationInfo As New System.Net.NetworkCredential("web.master@ossperu.com", "P@ssw0rd")
                oMessage.From = New System.Net.Mail.MailAddress("web.master@ossperu.com", "Web Master")
                basicAuthenticationInfo.UserName = "web.master@ossperu.com"
                'basicAuthenticationInfo.Password = "P@ssw0rd"
                'basicAuthenticationInfo.Password = "P@ssw0rd==" 
                'basicAuthenticationInfo.Password = "Password123$"
                basicAuthenticationInfo.Password = "J0e_S@tr1ani"
                oemailClient.Credentials = basicAuthenticationInfo
                oemailClient.Host = "mail.ossperu.com"
            ElseIf False Then 'No Funciona SSL / Si Funciona con TLS
                'oMessage.From = New System.Net.Mail.MailAddress("elmer.espinoza@gmail.com", "Elmer Espinoza")
                'basicAuthenticationInfo.UserName = "elmer.espinoza@gmail.com"
                'basicAuthenticationInfo.Password = "*****"

                'https://myaccount.google.com/u/0/lesssecureapps?pli=1

                oMessage.From = New System.Net.Mail.MailAddress("shirley.barreda.4269@gmail.com", "Shirley Barreda")
                oemailClient.Credentials = New System.Net.NetworkCredential("shirley.barreda.4269@gmail.com", "M@traquer4")

                oemailClient.EnableSsl = True
                oemailClient.Host = "smtp.gmail.com"
                oemailClient.Port = 587 '465

            Else
                oMessage.From = New System.Net.Mail.MailAddress("web.master@xerox.com", "Web Master")
                oemailClient.Host = "forwarder.mail.xerox.com"
            End If

            oMessage.To.Add(New System.Net.Mail.MailAddress(Destination))
            'oMessage.Bcc.Add(New System.Net.Mail.MailAddress("extranetperu@gmail.com", "Extranet"))

            If Trim(FileName) <> Nothing Then
                File.WriteAllBytes(System.IO.Path.GetTempPath + System.IO.Path.GetFileName(FileName), Convert.FromBase64String(FileBase64))
                oMessage.Attachments.Add(New System.Net.Mail.Attachment(System.IO.Path.GetTempPath + System.IO.Path.GetFileName(FileName)))
            End If

            oMessage.Body = oMessage.Body ' + "<br></br>CODEBASE64:<br>" + FileBase64.ToString

            oemailClient.Send(oMessage)

        Catch ex As Exception
            Status = ex.Message 'False
        End Try

        Return Status
    End Function


    <WebMethod()>
    Public Function GoogleMaps() As String
        'No funciona con STATIC MAP API necesita KEY
        'https://maps.googleapis.com/maps/api/staticmap?center=" + Latitud.ToString + "," + Longitud.ToString + "&zoom=16&size=600x600&maptype=roadmap&markers=icon:http://www.ossperu.com/gps.png%7Clabel:X%7C" + Latitud.ToString + "," + Longitud.ToString
        'https://www.google.com/maps/@-12.081796,-77.076810,18z
        'https://maps.google.com/?q=-12.081796,-77.07681&z=20 
        'https://www.google.com/maps/embed/v1/place?q=-12.081796,-77.07681&key=AIzaSyALP3TAMkGvSWNk9dKvPwR0vQ-0t526w5A
    End Function

    <WebMethod()>
    Public Function SendMailOss(ByVal Destination As String, ByVal Body As String) As String 'Boolean

        ' No Funciona requiere SSL 
        Dim msg As String = "Mail was Sent from Oss."

        'Try

        Dim Mail As New System.Net.Mail.MailMessage
        Dim SMTP As New System.Net.Mail.SmtpClient("mail.ossperu.com")

        Mail.Subject = "Mail from Oss"

        'https://myaccount.google.com/u/0/lesssecureapps?pli=1

        Mail.From = New System.Net.Mail.MailAddress("web.master@ossperu.com", "Web Master")

        SMTP.Credentials = New System.Net.NetworkCredential("web.master@ossperu.com", "J0e_S@tr1ani")

        Mail.To.Add(Destination) 'I used ByVal here for address

        Mail.Body = Body 'Message Here

        SMTP.EnableSsl = False
        SMTP.Port = 25 '587 '465
        SMTP.Send(Mail)

        'Catch ex As Exception

        'msg = ex.Message

        'End Try

        Return msg

    End Function


    <WebMethod()>
    Public Function SendMailGoogle(ByVal Destination As String, ByVal Body As String) As String 'Boolean

        ' No Funciona requiere SSL 
        Dim msg As String = "Mail was Sent from Google."

        'Try

        Dim Mail As New System.Net.Mail.MailMessage
        Dim SMTP As New System.Net.Mail.SmtpClient("smtp.gmail.com")

        Mail.Subject = "Mail from Google"

        'https://myaccount.google.com/u/0/lesssecureapps?pli=1

        Mail.From = New System.Net.Mail.MailAddress("shirley.barreda.4269@gmail.com", "Shirley Barreda")
        SMTP.Credentials = New System.Net.NetworkCredential("shirley.barreda.4269@gmail.com", "M@traquer4")

        Mail.To.Add(Destination) 'I used ByVal here for address

        Mail.Body = Body 'Message Here

        SMTP.EnableSsl = True
        SMTP.Port = 587 '465
        SMTP.Send(Mail)

        'Catch ex As Exception

        'msg = ex.Message

        'End Try

        Return msg

    End Function


    <WebMethod()> Function EncodeBase64(ByVal Text As String) As String
        Dim bytesToEncode As Byte()
        Dim encodedText As String
        bytesToEncode = Encoding.UTF8.GetBytes(Text)
        encodedText = Convert.ToBase64String(bytesToEncode)
        Return encodedText
    End Function


    <WebMethod()>
    Function DecodeBase64(ByVal Text As String) As String
        Dim decodedBytes As Byte()
        Dim decodedText As String
        decodedBytes = Convert.FromBase64String(Text)
        decodedText = Encoding.UTF8.GetString(decodedBytes)
        Return decodedText
    End Function

End Class


