Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Interfaces
Imports System.Configuration
Imports System.Data.SqlClient

Public Class MenuDC
    Inherits MarshalByRefObject

    Implements IMenu

    Public Function GetMenuStructure() As DataSet _
    Implements IMenu.GetMenuStructure
        Dim strCN As String = RegReader.getConnString
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "get_menu_structure"
        End With

        da.Fill(ds)

        cmd = Nothing
        cn.Close()

        Return ds
    End Function
End Class

