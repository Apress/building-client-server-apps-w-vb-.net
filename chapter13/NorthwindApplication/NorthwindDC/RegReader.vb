Option Explicit On 
Option Strict On

Imports Microsoft.Win32
Imports System.Configuration

Public Class RegReader
    Private Shared UserName As String
    Private Shared Password As String

    Shared Sub New()
        Dim regKey As RegistryKey

        regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\NorthwindTraders\\dbLogon", False)
        UserName = CType(regKey.GetValue("username"), String)
        Password = CType(regKey.GetValue("password"), String)
    End Sub

    Public Shared Function getConnString() As String
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim strID As String = "uid=" & UserName & ";pwd=" & Password & ";"

        Return strID & strCN
    End Function
End Class