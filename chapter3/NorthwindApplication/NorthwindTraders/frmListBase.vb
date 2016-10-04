Option Strict On
Option Explicit On 

Imports System.Drawing
Imports System.Drawing.Printing

Public Class frmListBase
    Inherits System.Windows.Forms.Form

    Private lvwColumnSorter As ListViewColumnSorter

    Private WithEvents mobjPD As PrintDocument
    Private mstrTitle As String
    Private mintPrintCount As Integer

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        lvwColumnSorter = New ListViewColumnSorter
        Me.lvwList.ListViewItemSorter = lvwColumnSorter
    End Sub

    Public Sub New(ByVal ReportTitle As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mstrTitle = ReportTitle
        lvwColumnSorter = New ListViewColumnSorter
        Me.lvwList.ListViewItemSorter = lvwColumnSorter
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboColumn As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents lvwList As System.Windows.Forms.ListView
    Friend WithEvents lblRecordCount As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents ctmListBase As System.Windows.Forms.ContextMenu
    Friend WithEvents ctmEdit As System.Windows.Forms.MenuItem
    Friend WithEvents ctmDelete As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmListBase))
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboColumn = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnFind = New System.Windows.Forms.Button
        Me.txtSearch = New System.Windows.Forms.TextBox
        Me.lvwList = New System.Windows.Forms.ListView
        Me.ctmListBase = New System.Windows.Forms.ContextMenu
        Me.ctmEdit = New System.Windows.Forms.MenuItem
        Me.ctmDelete = New System.Windows.Forms.MenuItem
        Me.lblRecordCount = New System.Windows.Forms.Label
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Column:"
        '
        'cboColumn
        '
        Me.cboColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboColumn.Location = New System.Drawing.Point(56, 8)
        Me.cboColumn.Name = "cboColumn"
        Me.cboColumn.Size = New System.Drawing.Size(121, 21)
        Me.cboColumn.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(192, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Search For:"
        '
        'btnFind
        '
        Me.btnFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFind.Image = CType(resources.GetObject("btnFind.Image"), System.Drawing.Image)
        Me.btnFind.Location = New System.Drawing.Point(408, 8)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(24, 23)
        Me.btnFind.TabIndex = 3
        '
        'txtSearch
        '
        Me.txtSearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSearch.Location = New System.Drawing.Point(256, 8)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(144, 20)
        Me.txtSearch.TabIndex = 4
        Me.txtSearch.Text = ""
        '
        'lvwList
        '
        Me.lvwList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvwList.ContextMenu = Me.ctmListBase
        Me.lvwList.FullRowSelect = True
        Me.lvwList.Location = New System.Drawing.Point(8, 40)
        Me.lvwList.MultiSelect = False
        Me.lvwList.Name = "lvwList"
        Me.lvwList.Size = New System.Drawing.Size(424, 192)
        Me.lvwList.TabIndex = 5
        Me.lvwList.View = System.Windows.Forms.View.Details
        '
        'ctmListBase
        '
        Me.ctmListBase.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.ctmEdit, Me.ctmDelete})
        '
        'ctmEdit
        '
        Me.ctmEdit.Index = 0
        Me.ctmEdit.Text = "&Edit"
        '
        'ctmDelete
        '
        Me.ctmDelete.Index = 1
        Me.ctmDelete.Text = "&Delete"
        '
        'lblRecordCount
        '
        Me.lblRecordCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblRecordCount.Location = New System.Drawing.Point(8, 232)
        Me.lblRecordCount.Name = "lblRecordCount"
        Me.lblRecordCount.Size = New System.Drawing.Size(100, 16)
        Me.lblRecordCount.TabIndex = 6
        Me.lblRecordCount.Text = "Record Count:"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(8, 264)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 7
        Me.btnDelete.Text = "&Delete"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(360, 264)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "&Close"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEdit.Location = New System.Drawing.Point(272, 264)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.TabIndex = 9
        Me.btnEdit.Text = "&Edit"
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(96, 264)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.TabIndex = 10
        Me.btnPrint.Text = "&Print"
        '
        'btnAdd
        '
        Me.btnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAdd.Location = New System.Drawing.Point(184, 264)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 11
        Me.btnAdd.Text = "&Add"
        '
        'frmListBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(440, 293)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.lblRecordCount)
        Me.Controls.Add(Me.lvwList)
        Me.Controls.Add(Me.txtSearch)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboColumn)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.Name = "frmListBase"
        Me.Text = "frmListBase"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmListBase_Paint(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Dim i As Integer
        'Resize the columns in the listview
        For i = 0 To lvwList.Columns.Count - 1
            lvwList.Columns.Item(i).Width = CInt((lvwList.Size.Width / _
            lvwList.Columns.Count) - 6)
        Next
        'Clear the line that we drew previously
        Me.CreateGraphics.Clear(Me.BackColor)

        'Draws the line above the buttons
        Me.CreateGraphics.DrawLine(New Pen(Color.Black, 1), _
                       lvwList.Location.X, btnDelete.Location.Y - 10, _
                       lvwList.Location.X + lvwList.Size.Width, _
                       btnDelete.Location.Y - 10)
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Protected Overridable Sub AddButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnAdd.Click
    End Sub

    Protected Overridable Sub DeleteButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnDelete.Click, ctmDelete.Click
    End Sub

    Protected Overridable Sub EditButton_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnEdit.Click, ctmEdit.Click, _
    lvwList.DoubleClick
    End Sub

    Private Sub frmListBase_Load(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer

        For i = 0 To lvwList.Columns.Count - 1
            cboColumn.Items.Add(lvwList.Columns(i).Text)
        Next
        If cboColumn.Items.Count > 0 Then
            cboColumn.SelectedIndex = 0
        End If
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnFind.Click
        Dim i As Integer
        Dim lst As ListViewItem

        i = cboColumn.SelectedIndex

        For Each lst In lvwList.Items
            If lst.SubItems(i).Text.ToUpper.StartsWith(txtSearch.Text.ToUpper) Then
                lst.Selected = True
                lst.EnsureVisible()
                Exit For
            End If
        Next

        lvwList.Focus()
    End Sub

    Private Sub lvwList_ColumnClick(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.ColumnClickEventArgs) _
    Handles lvwList.ColumnClick
        If (e.Column = lvwColumnSorter.SortColumn) Then
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If

        Me.lvwList.Sort()
    End Sub

    Protected Overridable Sub btnPrint_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim PPD As PrintPreviewDialog = New PrintPreviewDialog

        mobjPD = New PrintDocument
        PPD.Document = mobjPD
        PPD.WindowState = FormWindowState.Maximized
        PPD.PrintPreviewControl.AutoZoom = False
        PPD.PrintPreviewControl.Zoom = 1.0
        PPD.ShowDialog()
    End Sub

    Private Sub mPD_PrintPage(ByVal sender As Object, ByVal e As _
    System.Drawing.Printing.PrintPageEventArgs) Handles mobjPD.PrintPage
        Try
            Dim printFont As New Font("Arial", 10)
            Dim titleFont As New Font("Arial", 12, CType(FontStyle.Underline + _
  FontStyle.Bold, FontStyle), GraphicsUnit.Pixel)
            Dim headerFont As New Font("Arial", 24, FontStyle.Bold, _
  GraphicsUnit.Pixel)
            Dim headerLineHeight As Single = headerFont.GetHeight(e.Graphics)
            Dim lineHeight As Single = printFont.GetHeight(e.Graphics)
            Dim lPos As Single = e.MarginBounds.Left
            Dim yPos As Single = e.MarginBounds.Top
            Dim intLength As Integer
            Dim i, j As Integer

            'Print the header
            e.Graphics.DrawString("List Report", headerFont, Brushes.Black, _
  lPos, yPos)
            yPos += headerLineHeight
            lPos += 3

            'Print the Maintenance report type
            e.Graphics.DrawString(mstrTitle, printFont, Brushes.Black, lPos, yPos)

            yPos += lineHeight * 2

            'Reset the left margin
            lPos = e.MarginBounds.Left

            'Print the header columns
            For i = 0 To lvwList.Columns.Count - 1
                e.Graphics.DrawString(lvwList.Columns(i).Text, titleFont, _
 Brushes.Black, lPos, yPos, New StringFormat)
                If i < lvwList.Columns.Count - 1 Then
                    lPos += 150
                End If
            Next

            'Print the data to the report
            For i = mintPrintCount To lvwList.Items.Count - 1
                yPos += lineHeight
                lPos = e.MarginBounds.Left
                If lvwList.Items(i).Text.Length > 20 Then
                    intLength = 20
                Else
                    intLength = lvwList.Items(i).Text.Length
                End If
                e.Graphics.DrawString(lvwList.Items(i).Text.Substring(0, intLength), _
      printFont, Brushes.Black, lPos, yPos, New StringFormat)
                For j = 1 To lvwList.Columns.Count - 1
                    lPos += 150
                    If lvwList.Items(i).SubItems(j).Text.Length > 20 Then
                        intLength = 20
                    Else
                        intLength = lvwList.Items(i).SubItems(j).Text.Length
                    End If
                    e.Graphics.DrawString( _
     lvwList.Items(i).SubItems(j).Text.Substring(0, intLength), printFont, _
     Brushes.Black, lPos, yPos, New StringFormat)
                Next

                'If there are more pages, continue
                If yPos > e.MarginBounds.Bottom Then
                    e.HasMorePages = True
                    Exit For
                Else
                    mintPrintCount += 1
                End If
            Next
        Catch exc As Exception
            MessageBox.Show(exc.Message, "Print Error", MessageBoxButtons.OK, _
  MessageBoxIcon.Error)
        End Try
    End Sub

End Class
