Option Explicit On 
Option Strict On

Imports System.Xml.Serialization
Imports System.IO
Imports System.Web

Module Module1

    Sub Main()
        Console.WriteLine("Press a key to start the test.")
        Console.ReadLine()
        Console.WriteLine("Throwing an error...")
        Try
            Throw New ApplicationException("My error")
        Catch exc As Exception
            Dim objEL As LogErrorEvent
            Dim errArray() As structLoggedError

            objEL = objEL.getInstance

            objEL.LogErr(exc)
            Console.WriteLine("Error has been logged. Press any key to " _
            & "continue.")
            Console.ReadLine()
            Console.WriteLine("Reading error log...")
            errArray = objEL.RetrieveErrors()
            Console.WriteLine("Errors Retrieved. Press any key to continue.")
            Console.ReadLine()
            Console.WriteLine("Sending Errors...")
            objEL.SendErrors(errArray)
            Console.WriteLine("Errors Sent. Press Any Key To Exit Application.")
            Console.ReadLine()
        End Try
    End Sub


End Module


<Serializable()> Public Structure structLoggedError
    Public UserName As String
    Public ErrorTime As Date
    Public Source As String
    Public Message As String
    Public StackTrace As String
End Structure

Public Class CustomErrors
    Public ErrArray() As structLoggedError
End Class

Public Class LogErrorEvent
    Private Shared mobjLogErrEvent As LogErrorEvent

    Public Event ErrorLogged()
    Public Event ErrorsCleared()

    Private Const EVENTLOGNAME As String = "Northwind"
    Private Const ERRORFILENAME As String = "NorthwindErrors.xml"

    Public Shared Function getInstance() As LogErrorEvent
        If mobjLogErrEvent Is Nothing Then
            mobjLogErrEvent = New LogErrorEvent
        End If
        Return mobjLogErrEvent
    End Function

    Protected Sub New()
        Dim objEventLog As New EventLog

        If Not objEventLog.Exists(EVENTLOGNAME) Then
            objEventLog.CreateEventSource(EVENTLOGNAME, EVENTLOGNAME)
        End If
    End Sub

    Public Sub LogErr(ByVal exc As Exception)
        Dim objEventLog As New EventLog
        Dim objErr As structLoggedError
        Dim xmlSer As New XmlSerializer(GetType(structLoggedError))
        Dim strB As New System.Text.StringBuilder
        Dim txtWriter As New StringWriter(strB)
        Dim strText As String

        With objErr
            .UserName = System.Security.Principal.WindowsIdentity.GetCurrent.Name
            .ErrorTime = Now
            .Source = exc.Source
            .Message = exc.Message
            .StackTrace = exc.StackTrace
        End With

        xmlSer.Serialize(txtWriter, objErr)
        strText = strB.ToString

        objEventLog.WriteEntry(EVENTLOGNAME, strText, EventLogEntryType.Error)
        objEventLog = Nothing

        RaiseEvent ErrorLogged()
    End Sub

    Public Function RetrieveErrors() As structLoggedError()
        Dim objEL As New EventLog(EVENTLOGNAME)
        Dim objEntry As EventLogEntry
        Dim errArray(objEL.Entries.Count - 1) As structLoggedError
        Dim i As Integer = 0
        Dim xmlSer As New XmlSerializer(GetType(structLoggedError))

        For Each objEntry In objEL.Entries
            Dim txtReader As New StringReader(objEntry.Message)
            errArray(i) = CType(xmlSer.Deserialize(txtReader), structLoggedError)
            i += 1
            txtReader = Nothing
        Next
        objEL = Nothing
        Return errArray
    End Function

    Public Sub SendErrors(ByVal errArray() As structLoggedError)
        Dim objAllErrors As New CustomErrors
        Dim xmlFileSer As New XmlSerializer(GetType(CustomErrors))
        Dim dirPath As Environment
        Dim strPath As String = _
             dirPath.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim fStream As New FileStream(strPath & "\" & ERRORFILENAME, _
             FileMode.Create)
        Dim email As New Mail.MailMessage
        Dim emailAttach As Mail.MailAttachment

        objAllErrors.ErrArray = errArray
        xmlFileSer.Serialize(fStream, objAllErrors)
        fStream.Close()
        xmlFileSer = Nothing
        emailAttach = New Mail.MailAttachment(strPath & "\" & ERRORFILENAME)
        With email
            .From = "youremail@somewhere.com"
            .To = "youremail@somewhere.com"
            .Subject = "Northwind Error Log"
            .Priority = Web.Mail.MailPriority.High
            .Attachments.Add(emailAttach)
        End With
        Mail.SmtpMail.Send(email)
        Kill(strPath & "\" & ERRORFILENAME)
        ClearErrors()
    End Sub

    Public Sub ClearErrors()
        Dim objEventLog As New EventLog(EVENTLOGNAME)
        objEventLog.Clear()
        objEventLog = Nothing
        RaiseEvent ErrorsCleared()
    End Sub

End Class


