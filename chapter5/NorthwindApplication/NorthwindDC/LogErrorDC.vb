Option Strict On
Option Explicit On 

Imports System.Data.SqlClient
Imports System.Configuration
Imports NorthwindTraders.NorthwindShared.Interfaces

Public Class LogErrorDC
    Inherits MarshalByRefObject

    Implements ILogError

    Public Sub Save(ByVal exc As Exception) Implements ILogError.Save
        Dim strCN As String = _
ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_application_errors_save"
            .Parameters.Add("@user_name", _
            System.Security.Principal.WindowsIdentity.GetCurrent.Name)
            .Parameters.Add("@error_message", exc.Message)
            .Parameters.Add("@error_time", Now)
            If exc.StackTrace = "" Then
                .Parameters.Add("@stack_trace", DBNull.Value)
            Else
                .Parameters.Add("@stack_trace", exc.StackTrace)
            End If
        End With

        cmd.ExecuteNonQuery()
        cmd = Nothing
        cn.Close()
        cn = Nothing
    End Sub
End Class

