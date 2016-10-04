Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindUC

Public Class frmEmployeeList
    Inherits NorthwindTraders.frmListBase

    Private mobjEmployeeMgr As EmployeeMgr
    Private WithEvents mfrmEdit As frmEmployeeEdit
    Private WithEvents mobjEmployee As Employee
    Private mobjTerritoryMgr As TerritoryMgr

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New("Employees")

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
        'frmEmployeeList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(440, 293)
        Me.Name = "frmEmployeeList"
        Me.Text = "Employee List"

    End Sub

#End Region

    Private Sub LoadList()
        Dim objEmployee As Employee
        Dim objDictEnt As DictionaryEntry

        Try
            mobjTerritoryMgr = New TerritoryMgr(True)

            lvwList.BeginUpdate()

            If lvwList.Columns.Count = 0 Then
                With lvwList
                    .Columns.Add("Last Name", CInt(.Size.Width / 3) - _
                    8, HorizontalAlignment.Left)
                    .Columns.Add("First Name", CInt(.Size.Width / 3) - _
                    8, HorizontalAlignment.Left)
                    .Columns.Add("Title", CInt(.Size.Width / 3) - 8, _
                    HorizontalAlignment.Left)
                End With
            End If

            lvwList.Items.Clear()

            mobjEmployeeMgr = mobjEmployeeMgr.GetInstance

            For Each objDictEnt In mobjEmployeeMgr
                objEmployee = CType(objDictEnt.Value, Employee)
                Dim lst As New ListViewItem(objEmployee.LastName)
                lst.Tag = objEmployee.EmployeeID
                lst.SubItems.Add(objEmployee.FirstName)
                lst.SubItems.Add(objEmployee.Title)
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
                mobjEmployee = New Employee
                mfrmEdit = New frmEmployeeEdit(mobjEmployee, _
                mobjTerritoryMgr)
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

    Private Sub mobjEmployee_ObjectChanged(ByVal sender As Object, _
    ByVal e As ChangedEventArgs) _
    Handles mobjEmployee.ObjectChanged
        Try
            Dim lst As ListViewItem
            Dim objEmployee As Employee = CType(sender, Employee)

            Select Case e.Change
                Case ChangedEventArgs.eChange.Added
                    mobjEmployeeMgr.Add(objEmployee)

                    lst = New ListViewItem(objEmployee.LastName)
                    lst.Tag = objEmployee.EmployeeID
                    lst.SubItems.Add(objEmployee.FirstName)
                    lst.SubItems.Add(objEmployee.Title)
                    lvwList.Items.Add(lst)
                    lblRecordCount.Text = "Record Count: " & _
                    lvwList.Items.Count
                Case ChangedEventArgs.eChange.Updated
                    For Each lst In lvwList.Items
                        If Convert.ToInt32(lst.Tag) = _
                        objEmployee.EmployeeID Then
                            lst.Text = objEmployee.LastName
                            lst.SubItems(1).Text = objEmployee.FirstName
                            lst.SubItems(2).Text = objEmployee.Title
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
                    Cursor = Cursors.WaitCursor
                    mobjEmployee = _
                    mobjEmployeeMgr.Item(lvwList.SelectedItems(0).Tag)
                    mobjEmployee.LoadRecord(mobjTerritoryMgr)
                    mfrmEdit = New frmEmployeeEdit(mobjEmployee, _
                    mobjTerritoryMgr)
                    mfrmEdit.MdiParent = Me.MdiParent
                    mfrmEdit.Show()
                End If
            End If
        Catch exc As Exception
            LogException(exc)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Protected Overrides Sub DeleteButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs)
        Dim objEmployee As Employee
        Dim dlgResult As DialogResult

        Try
            If lvwList.SelectedItems.Count > 0 Then
                objEmployee = _
                mobjEmployeeMgr.Item(lvwList.SelectedItems(0).Tag)
                dlgResult = MessageBox.Show("Do you want to delete " _
                & "employee: " & objEmployee.ToString & "?", _
                "Confirm Delete", MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question)
                If dlgResult = DialogResult.Yes Then
                    objEmployee.Delete()
                    mobjEmployeeMgr.Remove(objEmployee.EmployeeID)
                    lvwList.SelectedItems(0).Remove()
                    lblRecordCount.Text = "Record Count: " _
                    & lvwList.Items.Count
                End If
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub
End Class
