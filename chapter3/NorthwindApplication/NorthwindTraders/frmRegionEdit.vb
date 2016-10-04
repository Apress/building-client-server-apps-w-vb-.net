Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindUC

Public Class frmRegionEdit
    Inherits NorthwindTraders.frmEditBase

    Private mobjRegion As Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef objRegion As Region)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mobjRegion = objRegion
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
    Friend WithEvents txtRegionDescription As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtRegionDescription = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(178, 50)
        Me.btnOK.Name = "btnOK"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(266, 50)
        Me.btnCancel.Name = "btnCancel"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Region Description:"
        '
        'txtRegionDescription
        '
        Me.txtRegionDescription.Location = New System.Drawing.Point(128, 16)
        Me.txtRegionDescription.Name = "txtRegionDescription"
        Me.txtRegionDescription.Size = New System.Drawing.Size(208, 20)
        Me.txtRegionDescription.TabIndex = 3
        Me.txtRegionDescription.Text = ""
        '
        'frmRegionEdit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(346, 79)
        Me.Controls.Add(Me.txtRegionDescription)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmRegionEdit"
        Me.Text = "Region [Detail]"
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.txtRegionDescription, 0)
        Me.Controls.SetChildIndex(Me.btnOK, 0)
        Me.Controls.SetChildIndex(Me.btnCancel, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmRegionEdit_Load(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        If mobjRegion.RegionID > 0 Then
            txtRegionDescription.Text = mobjRegion.RegionDescription.Trim
        End If
    End Sub

    Private Sub txtRegionDescription_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtRegionDescription.Validated
        Dim txt As TextBox

        txt = CType(sender, TextBox)
        mobjRegion.RegionDescription = txt.Text
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnOK.Click
        If mobjRegion.IsDirty Then
            mobjRegion.Save()
        End If
        Close()
    End Sub

End Class
