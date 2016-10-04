Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindUC

Public Class frmTerritoryList
    Inherits NorthwindTraders.frmListBase

    Private mobjTerritoryMgr As TerritoryMgr
    Private WithEvents mfrmEdit As frmTerritoryEdit
    Private WithEvents mobjTerritory As Territory

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New("Territories")

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        LoadList()
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub LoadList()
        Dim objTerritory As Territory
        Dim objDictEnt As DictionaryEntry

        Try
            lvwList.BeginUpdate()

            If lvwList.Columns.Count = 0 Then
                With lvwList
                    .Columns.Add("Territory ID", CInt(.Size.Width / 1) - 8, _
                    HorizontalAlignment.Left)
                    .Columns.Add("Territory Description", CInt(.Size.Width _
                    / 1) - 8, HorizontalAlignment.Left)
                    .Columns.Add("Region", CInt(.Size.Width / 1) - 8, _
                    HorizontalAlignment.Left)
                End With
            End If

            lvwList.Items.Clear()

            mobjTerritoryMgr = New TerritoryMgr

            For Each objDictEnt In mobjTerritoryMgr
                objTerritory = CType(objDictEnt.Value, Territory)
                Dim lst As New ListViewItem(objTerritory.TerritoryID)
                lst.SubItems.Add(objTerritory.TerritoryDescription)
                lst.SubItems.Add(objTerritory.Region.RegionDescription)
                lvwList.Items.Add(lst)
            Next

            lvwList.EndUpdate()

            lblRecordCount.Text = "Record Count: " & lvwList.Items.Count
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Protected Overrides Sub AddButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs)
        Try
            If mfrmEdit Is Nothing Then
                mobjTerritory = New Territory
                mfrmEdit = New frmTerritoryEdit(mobjTerritory)
                mfrmEdit.MdiParent = Me.MdiParent
                mfrmEdit.Show()
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub mfrmEdit_Closed(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles mfrmEdit.Closed
        mfrmEdit = Nothing
    End Sub

    Private Sub mobjTerritory_ObjectChanged(ByVal sender As Object, _
    ByVal e As ChangedEventArgs) _
    Handles mobjTerritory.ObjectChanged
        Try
            Dim lst As ListViewItem
            Dim objTerritory As Territory = CType(sender, Territory)

            Select Case e.Change
                Case ChangedEventArgs.eChange.Added
                    mobjTerritoryMgr.Add(objTerritory)
                    lst = New ListViewItem(objTerritory.TerritoryID)
                    lst.SubItems.Add(objTerritory.TerritoryDescription)
                    lst.SubItems.Add(objTerritory.Region.RegionDescription)
                    lvwList.Items.Add(lst)
                    lblRecordCount.Text = "Record Count: " _
                    & lvwList.Items.Count
                Case ChangedEventArgs.eChange.Updated
                    For Each lst In lvwList.Items
                        If lst.Text = objTerritory.TerritoryID Then
                            lst.SubItems(1).Text = _
                            objTerritory.TerritoryDescription
                            lst.SubItems(2).Text = _
                            objTerritory.Region.RegionDescription
                            Exit For
                        End If
                    Next
            End Select

            lvwList.Sort()
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Protected Overrides Sub EditButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs)
        Try
            If mfrmEdit Is Nothing Then
                If lvwList.SelectedItems.Count > 0 Then
                    mobjTerritory = _
                    mobjTerritoryMgr.Item(lvwList.SelectedItems(0).Text)
                    mobjTerritory.LoadRecord()
                    mfrmEdit = New frmTerritoryEdit(mobjTerritory)
                    mfrmEdit.MdiParent = Me.MdiParent
                    mfrmEdit.Show()
                End If
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Protected Overrides Sub DeleteButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs)
        Try
            Dim objTerritory As Territory
            objTerritory = mobjTerritoryMgr.Item(lvwList.SelectedItems(0).Text)
            objTerritory.Delete()
            mobjTerritoryMgr.Remove(objTerritory.TerritoryID)
            lvwList.SelectedItems(0).Remove()
            lblRecordCount.Text = "Record Count: " & lvwList.Items.Count
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

End Class
