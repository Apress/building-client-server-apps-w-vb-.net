Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindShared.Interfaces
Imports NorthwindTraders.NorthwindShared.Structures
Imports System.Configuration
Imports System.Data.SqlClient

Public Class RegionDC
    Inherits MarshalByRefObject

    Implements IRegion

#Region " Private Attributes"
    Private mintRegionID As Integer
    Private mstrRegionDescription As String
#End Region

#Region " Public Attributes"
    Public ReadOnly Property RegionID() As Integer
        Get
            Return mintRegionID
        End Get
    End Property

    Public Property RegionDescription() As String
        Get
            Return mstrRegionDescription
        End Get
        Set(ByVal Value As String)
            mstrRegionDescription = Value
        End Set
    End Property
#End Region

    Public Sub Delete(ByVal intID As Integer) Implements NorthwindShared.Interfaces.IRegion.Delete
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_region_delete"
            .Parameters.Add("@id", intID)
            .ExecuteNonQuery()
        End With

        cmd = Nothing
        cn.Close()
    End Sub

    Public Function LoadProxy() As System.Data.DataSet Implements NorthwindShared.Interfaces.IRegion.LoadProxy
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_region_getall"
        End With

        da.Fill(ds)

        cmd = Nothing
        cn.Close()

        Return ds
    End Function

    Public Function LoadRecord(ByVal intID As Integer) As NorthwindShared.Structures.structRegion Implements NorthwindShared.Interfaces.IRegion.LoadRecord
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        Dim sRegion As structRegion

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_region_getone"
            .Parameters.Add("@id", intID)
        End With

        da.Fill(ds)

        cmd = Nothing
        cn.Close()

        With ds.Tables(0).Rows(0)
            sRegion.RegionID = Convert.ToInt32(.Item("RegionID"))
            sRegion.RegionDescription = _
               Convert.ToString(.Item("RegionDescription"))
        End With
        ds = Nothing

        Return sRegion
    End Function

    Public Sub Save(ByVal sRegion As NorthwindShared.Structures.structRegion, ByRef intID As Integer) Implements NorthwindShared.Interfaces.IRegion.Save
        Dim strCN As String = ConfigurationSettings.AppSettings("Northwind_DSN")
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand

        With sRegion
            Me.mintRegionID = .RegionID
            Me.RegionDescription = .RegionDescription
        End With

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_region_save"
            .Parameters.Add("@id", mintRegionID)
            .Parameters.Add("@region", mstrRegionDescription)
            .Parameters.Add("@new_id", SqlDbType.Int).Direction = _
                ParameterDirection.Output
            .ExecuteNonQuery()
            intID = Convert.ToInt32(.Parameters.Item("@id").Value)
        End With

        cmd = Nothing
        cn.Close()
    End Sub
End Class

