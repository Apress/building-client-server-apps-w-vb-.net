Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindUC

Public Class frmRegionList
    Inherits NorthwindTraders.frmListBase

    Private mobjRegionMgr As RegionMgr

    Private WithEvents mfrmEdit As frmRegionEdit
    Private WithEvents mobjRegion As Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New("Regions")

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
        '
        'frmRegionList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(440, 293)
        Me.Name = "frmRegionList"
        Me.Text = "Region List"

    End Sub

#End Region

    Private Sub LoadList()
        Dim objRegion As Region
        Dim objDictEnt As DictionaryEntry

        lvwList.BeginUpdate()

        If lvwList.Columns.Count = 0 Then
            With lvwList
                .Columns.Add("Region Description", CInt(.Size.Width / 1) - 8, _
 HorizontalAlignment.Left)
            End With
        End If

        lvwList.Items.Clear()

        mobjRegionMgr = mobjRegionMgr.GetInstance

        For Each objDictEnt In mobjRegionMgr
            objRegion = CType(objDictEnt.Value, Region)
            Dim lst As New ListViewItem(objRegion.ToString)
            lst.Tag = objRegion.RegionID
            lvwList.Items.Add(lst)
        Next

        lvwList.EndUpdate()
        lblRecordCount.Text = "Record Count: " & lvwList.Items.Count

    End Sub

    Protected Overrides Sub AddButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs)
        If mfrmEdit Is Nothing Then
            Cursor = Cursors.WaitCursor
            mobjRegion = New Region
            mfrmEdit = New frmRegionEdit(mobjRegion)
            mfrmEdit.MdiParent = Me.MdiParent
            mfrmEdit.Show()
            Cursor = Cursors.Default
        End If
    End Sub

    Private Sub mfrmEdit_Closed(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles mfrmEdit.Closed
        mfrmEdit = Nothing
    End Sub

    Private Sub mobjRegion_ObjectChanged(ByVal sender As Object, _
    ByVal e As NorthwindTraders.NorthwindUC.ChangedEventArgs) Handles _
    mobjRegion.ObjectChanged
        Dim lst As ListViewItem
        Dim objRegion As Region = CType(sender, Region)

        Select Case e.Change
            Case ChangedEventArgs.eChange.Added
                mobjRegionMgr.Add(objRegion)
                lst = New ListViewItem(objRegion.RegionDescription)
                lst.Tag = objRegion.RegionID
                lvwList.Items.Add(lst)
                lblRecordCount.Text = "Record Count: " & lvwList.Items.Count
            Case ChangedEventArgs.eChange.Updated
                For Each lst In lvwList.Items
                    If Convert.ToInt32(lst.Tag) = objRegion.RegionID Then
                        lst.Text = objRegion.RegionDescription
                        Exit For
                    End If
                Next
        End Select
    End Sub

    Protected Overrides Sub EditButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs)
        If mfrmEdit Is Nothing Then
            If lvwList.SelectedItems.Count > 0 Then
                mobjRegion = mobjRegionMgr.Item(lvwList.SelectedItems(0).Tag)
                mobjRegion.LoadRecord()
                mfrmEdit = New frmRegionEdit(mobjRegion)
                mfrmEdit.MdiParent = Me.MdiParent
                mfrmEdit.Show()
            End If
        End If
    End Sub

    Protected Overrides Sub DeleteButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs)
        Dim objRegion As Region
        Dim dlgResult As DialogResult
        If lvwList.SelectedItems.Count > 0 Then
            objRegion = mobjRegionMgr.Item(lvwList.SelectedItems(0).Tag)
            dlgResult = MessageBox.Show("Do you want to delete the " _
            & objRegion.RegionDescription & " region?", _
            "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dlgResult = DialogResult.Yes Then
                objRegion.Delete()
                mobjRegionMgr.Remove(objRegion.RegionID)
                lvwList.SelectedItems(0).Remove()
                lblRecordCount.Text = "Record Count: " & lvwList.Items.Count
            End If
        End If
    End Sub

End Class
