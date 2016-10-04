Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindShared.Structures
Imports NorthwindTraders.NorthwindShared.Interfaces

Public Class Region

    Private mblnDirty As Boolean = False
    Public Loading As Boolean
    Private Const LISTENER As String = "RegionDC.rem"
    Private msRegion As structRegion

    Public Event ObjectChanged(ByVal sender As Object, _
        ByVal e As ChangedEventArgs)

#Region " Private Attributes"
    Private mintRegionID As Integer = 0
    Private mstrRegionDescription As String = ""
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
            If mstrRegionDescription.trim <> Value Then
                If Not Loading Then
                    mblnDirty = True
                End If
                mstrRegionDescription = Value
            End If
        End Set
    End Property

#End Region

    Public Sub New()
        'No code necessary right now
    End Sub

    Public Sub New(ByVal intID As Integer)
        mintRegionID = intID
    End Sub

    Public ReadOnly Property IsDirty() As Boolean
        Get
            Return mblnDirty
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return mstrRegionDescription
    End Function

    Public Sub LoadRecord()
        Dim objIRegion As IRegion

        objIRegion = CType(Activator.GetObject(GetType(IRegion), _
   AppConstants.REMOTEOBJECTS & LISTENER), IRegion)
        msRegion = objIRegion.LoadRecord(mintRegionID)
        objIRegion = Nothing

        LoadObject()
    End Sub

    Private Sub LoadObject()
        With msRegion
            Me.mintRegionID = .RegionID
            Me.mstrRegionDescription = .RegionDescription
        End With
    End Sub

    Public Sub Delete()
        Dim objIRegion As IRegion

        objIRegion = CType(Activator.GetObject(GetType(IRegion), _
    AppConstants.REMOTEOBJECTS & LISTENER), IRegion)
        objIRegion.Delete(mintRegionID)
        objIRegion = Nothing
    End Sub

    Public Sub Save()
        'Only save the object if it has been edited
        If mblnDirty = True Then
            Dim objIRegion As IRegion
            Dim intID As Integer
            Dim sRegion As structRegion

            'Store the original ID of the object
            intID = mintRegionID

            'Assign the values of the object to a structure of
            'type structRegion
            With sRegion
                .RegionID = mintRegionID
                .RegionDescription = mstrRegionDescription
            End With

            'Obtain a reference to the remote object
            objIRegion = CType(Activator.GetObject(GetType(IRegion), _
            AppConstants.REMOTEOBJECTS & LISTENER), IRegion)

            'Save the object
            objIRegion.Save(sRegion, mintRegionID)
            objIRegion = Nothing

            'Check the original ID and see if it was a new object
            If intID = 0 Then
                'If it was, raise the event as an added
                RaiseEvent ObjectChanged(Me, New _
                ChangedEventArgs(ChangedEventArgs.eChange.Added))
            Else
                'If it was not, raise the event as an update
                RaiseEvent ObjectChanged(Me, New _
                ChangedEventArgs(ChangedEventArgs.eChange.Updated))
            End If
            'The object is no longer dirty
            mblnDirty = False
        End If
    End Sub

End Class

Public Class RegionMgr
    Inherits System.Collections.DictionaryBase

    Private Shared mobjRegionMgr As RegionMgr

    Public Shared Function GetInstance() As RegionMgr
        If mobjRegionMgr Is Nothing Then
            mobjRegionMgr = New RegionMgr
        End If
        Return mobjRegionMgr
    End Function

    Protected Sub New()
        Load()
    End Sub

    Public Sub Add(ByVal obj As Region)
        dictionary.Add(obj.RegionID, obj)
    End Sub

    Public Function Item(ByVal Key As Object) As Region
        Return CType(dictionary.Item(Key), Region)
    End Function

    Public Sub Remove(ByVal Key As Object)
        dictionary.Remove(Key)
    End Sub

    Private Sub Load()
        Dim objIRegion As IRegion
        Dim dRow As DataRow
        Dim ds As DataSet

        objIRegion = CType(Activator.GetObject(GetType(IRegion), _
        AppConstants.REMOTEOBJECTS & "RegionDC.rem"), IRegion)
        ds = objIRegion.LoadProxy()
        objIRegion = Nothing

        For Each dRow In ds.Tables(0).Rows
            Dim objRegion As New Region(Convert.ToInt32(dRow.Item("RegionID")))
            With objRegion
                .Loading = True
                .RegionDescription = _
                     Convert.ToString(dRow.Item("RegionDescription")).Trim
                .Loading = False
            End With
            Me.Add(objRegion)
        Next

        ds = Nothing
    End Sub

    Public Sub Refresh()
        dictionary.Clear()
        Load()
    End Sub

End Class

