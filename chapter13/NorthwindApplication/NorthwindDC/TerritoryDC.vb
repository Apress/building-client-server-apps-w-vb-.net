Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Interfaces
Imports NorthwindTraders.NorthwindShared.Structures
Imports BusinessRules.Errors
Imports System.Configuration
Imports System.Data.SqlClient
Imports NorthwindTraders.NorthwindShared.Errors

Public Class TerritoryDC
    Inherits MarshalByRefObject

    Implements ITerritory

    Private mobjBusErr As BusinessErrors

#Region " Private Attributes"
    Private mstrTerritoryID As String
    Private mstrTerritoryDescription As String
    Private mintRegionID As Integer
#End Region

#Region " Public Attributes"

    Public Property TerritoryID() As String
        Get
            Return mstrTerritoryID
        End Get
        Set(ByVal Value As String)
            Try
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Territory ID")
                End If

                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                If Value.Length > 20 Then
                    Throw New MaximumLengthException(20)
                End If

                mstrTerritoryID = Value
            Catch exc As Exception
                mobjBusErr.Add("Territory ID", exc.Message)
            End Try
        End Set
    End Property
    Public Property TerritoryDescription() As String
        Get
            Return mstrTerritoryDescription
        End Get
        Set(ByVal Value As String)
            Try
                If Value Is Nothing Then
                    Throw New ArgumentNullException("Territory Description")
                End If

                If Value.Length = 0 Then
                    Throw New ZeroLengthException
                End If

                If Value.Length > 50 Then
                    Throw New MaximumLengthException(50)
                End If

                mstrTerritoryDescription = Value
            Catch exc As Exception
                mobjBusErr.Add("Territory Description", exc.Message)
            End Try
        End Set
    End Property

    Public Property RegionID() As Integer
        Get
            Return mintRegionID
        End Get
        Set(ByVal Value As Integer)
            Try
                If Value = 0 Then
                    Throw New ArgumentNullException("Region")
                End If

                mintRegionID = Value
            Catch exc As Exception
                mobjBusErr.Add("Region", exc.Message)
            End Try
        End Set
    End Property
#End Region

    Public Function LoadProxy() As DataSet Implements ITerritory.LoadProxy
        Dim strCN As String = RegReader.getConnString
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_territory_getall"
        End With

        da.Fill(ds)

        cmd = Nothing
        cn.Close()

        Return ds
    End Function

    Public Function LoadRecord(ByVal strID As String) As _
    structTerritory Implements ITerritory.LoadRecord
        Dim strCN As String = RegReader.getConnString
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        Dim sTerritory As structTerritory

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_territory_getone"
            .Parameters.Add("@id", strID)
        End With

        da.Fill(ds)

        cmd = Nothing
        cn.Close()

        With ds.Tables(0).Rows(0)
            sTerritory.TerritoryID = Convert.ToString(.Item("TerritoryID"))
            sTerritory.TerritoryDescription = Convert.ToString(.Item("TerritoryDescription"))
            sTerritory.RegionID = Convert.ToInt32(.Item("RegionID"))
        End With

        ds = Nothing

        Return sTerritory
    End Function

    Public Sub Delete(ByVal strID As String) Implements ITerritory.Delete
        Dim strCN As String = RegReader.getConnString
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand

        cn.Open()

        With cmd
            .Connection = cn
            .CommandType = CommandType.StoredProcedure
            .CommandText = "usp_territory_delete"
            .Parameters.Add("@id", strID)
            .ExecuteNonQuery()
        End With

        cmd = Nothing
        cn.Close()
    End Sub

    Public Function Save(ByVal sTerritory As structTerritory) _
    As BusinessErrors Implements ITerritory.Save
        Dim strCN As String = RegReader.getConnString
        Dim cn As New SqlConnection(strCN)
        Dim cmd As New SqlCommand
        Dim intNew As Integer

        mobjBusErr = New BusinessErrors

        With sTerritory
            Me.TerritoryID = .TerritoryID
            Me.TerritoryDescription = .TerritoryDescription
            Me.RegionID = .RegionID
        End With

        If mobjBusErr.Count = 0 Then

            If sTerritory.IsNew Then
                intNew = 1
            Else
                intNew = 0
            End If

            cn.Open()

            With cmd
                .Connection = cn
                .CommandType = CommandType.StoredProcedure
                .CommandText = "usp_territory_save"
                .Parameters.Add("@id", mstrTerritoryID)
                .Parameters.Add("@territory", mstrTerritoryDescription)
                .Parameters.Add("@region_id", mintRegionID)
                .Parameters.Add("@new", intNew)
                .ExecuteNonQuery()
            End With

            cmd = Nothing
            cn.Close()
        End If

        Return mobjBusErr
    End Function

    Public Function GetBusinessRules() As BusinessErrors _
    Implements ITerritory.GetBusinessRules
        Dim objBusRules As New BusinessErrors
        With objBusRules
            .Add("Territory ID", "The value cannot be null.")
            .Add("Territory ID", "The value cannot be more than 20 characters in length.")
            .Add("Territory Description", "The value cannot be null.")
            .Add("Territory Description", "The value cannot be more than 50 characters in length.")
            .Add("Region", "The value cannot be null.")
        End With

        Return objBusRules
    End Function
End Class

