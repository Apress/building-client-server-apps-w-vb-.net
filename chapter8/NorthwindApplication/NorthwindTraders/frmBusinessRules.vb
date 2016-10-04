Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindShared.Errors

Public Class frmBusinessRules
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal objBE As BusinessErrors)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        LoadListView(objBE)
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
    Friend WithEvents lvwList As System.Windows.Forms.ListView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBusinessRules))
        Me.lvwList = New System.Windows.Forms.ListView
        Me.btnClose = New System.Windows.Forms.Button
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.SuspendLayout()
        '
        'lvwList
        '
        Me.lvwList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvwList.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lvwList.Location = New System.Drawing.Point(8, 8)
        Me.lvwList.Name = "lvwList"
        Me.lvwList.Size = New System.Drawing.Size(280, 232)
        Me.lvwList.TabIndex = 0
        Me.lvwList.View = System.Windows.Forms.View.Details
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(216, 248)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "&Close"
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Field"
        Me.ColumnHeader1.Width = 120
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Rule"
        Me.ColumnHeader2.Width = 154
        '
        'frmBusinessRules
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lvwList)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmBusinessRules"
        Me.Text = "Business Rules"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub LoadListView(ByVal objBusErr As BusinessErrors)
        Dim i As Integer

        For i = 0 To objBusErr.Count - 1
            Dim lst As New ListViewItem(objBusErr.Item(i).errProperty)
            lst.SubItems.Add(objBusErr.Item(i).errMessage)
            lvwList.Items.Add(lst)
        Next
    End Sub

    Private Sub frmBusinessRules_Paint(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Dim i As Integer

        For i = 0 To lvwList.Columns.Count - 1
            lvwList.Columns.Item(i).Width = CInt((lvwList.Size.Width / _
            lvwList.Columns.Count) - 6)
        Next
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

End Class
