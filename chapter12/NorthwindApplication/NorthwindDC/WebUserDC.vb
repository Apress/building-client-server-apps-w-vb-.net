Option Explicit On 
Option Strict On

Imports System.Data.SqlClient
Imports NorthwindTraders.NorthwindShared.Interfaces

Public Class WebUserDC
    Inherits MarshalByRefObject

    Implements IWebUser

    Public Function ValidateUser(ByVal strUser As String, _
    ByVal strPass As String) As Boolean Implements IWebUser.ValidateUser
        Dim strCN As String = RegReader.getConnString
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim intValid As Integer

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_ValidateUser"
            .Parameters.Add("@uid", strUser)
            .Parameters.Add("@pwd", strPass)
            .Parameters.Add("@valid", Nothing).Direction = ParameterDirection.Output
            .ExecuteNonQuery()
            intValid = Convert.ToInt32(.Parameters("@valid").Value)
        End With

        cmd = Nothing
        cn.Close()

        If intValid = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AddUser(ByVal strUser As String, _
    ByVal strPass As String) As Boolean Implements IWebUser.AddUser
        Dim strCN As String = RegReader.getConnString
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim intValid As Integer

        Try
            cn.Open()

            With cmd
                .Connection = cn
                .CommandType = CommandType.StoredProcedure
                .CommandText = "usp_AddNewUser"
                .Parameters.Add("@uid", strUser)
                .Parameters.Add("@pwd", strPass)
                .ExecuteNonQuery()
            End With

            cmd = Nothing
            cn.Close()

            Return True
        Catch
            Return False
        End Try
    End Function

End Class
