Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindUC

Public Class frmRegionEdit
    Inherits NorthwindTraders.frmEditBase

    Private WithEvents mobjRegion As Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef objRegion As Region)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mobjRegion = objRegion

        If Not mobjRegion.IsValid Then
            btnOK.Enabled = False
        End If
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
    Friend WithEvents txtRegionDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblRegionDesc As System.Windows.Forms.Label

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtRegionDescription = New System.Windows.Forms.TextBox
        Me.lblRegionDesc = New System.Windows.Forms.Label
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
        'btnRules
        '
        Me.btnRules.Location = New System.Drawing.Point(8, 48)
        Me.btnRules.Name = "btnRules"
        '
        'txtRegionDescription
        '
        Me.erpMain.SetIconAlignment(Me.txtRegionDescription, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtRegionDescription.Location = New System.Drawing.Point(128, 16)
        Me.txtRegionDescription.Name = "txtRegionDescription"
        Me.txtRegionDescription.Size = New System.Drawing.Size(208, 20)
        Me.txtRegionDescription.TabIndex = 3
        Me.txtRegionDescription.Text = ""
        '
        'lblRegionDesc
        '
        Me.lblRegionDesc.Location = New System.Drawing.Point(8, 16)
        Me.lblRegionDesc.Name = "lblRegionDesc"
        Me.lblRegionDesc.Size = New System.Drawing.Size(104, 16)
        Me.lblRegionDesc.TabIndex = 4
        Me.lblRegionDesc.Text = "Region Description:"
        '
        'frmRegionEdit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(346, 79)
        Me.Controls.Add(Me.lblRegionDesc)
        Me.Controls.Add(Me.txtRegionDescription)
        Me.Name = "frmRegionEdit"
        Me.Text = "Region [Detail]"
        Me.Controls.SetChildIndex(Me.btnRules, 0)
        Me.Controls.SetChildIndex(Me.txtRegionDescription, 0)
        Me.Controls.SetChildIndex(Me.lblRegionDesc, 0)
        Me.Controls.SetChildIndex(Me.btnOK, 0)
        Me.Controls.SetChildIndex(Me.btnCancel, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub mobjRegion_BrokenRule(ByVal IsBroken As Boolean) _
    Handles mobjRegion.BrokenRule
        If IsBroken Then
            btnOK.Enabled = False
        Else
            btnOK.Enabled = True
        End If
    End Sub

    Private Sub frmRegionEdit_Load(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If mobjRegion.RegionID > 0 Then
                txtRegionDescription.Text = mobjRegion.RegionDescription.Trim
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub txtRegionDescription_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtRegionDescription.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjRegion.RegionDescription = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub frmRegionEdit_Closing(ByVal sender As Object, _
    ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim dlgResult As DialogResult

        Try
            If mobjRegion.IsDirty Then
                dlgResult = MessageBox.Show("The Region information has " _
                & "changed, do you want to exit without saving your changes?", _
                "Confirm Cancel", MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question)

                If dlgResult = DialogResult.No Then
                    e.Cancel = True
                Else
                    mobjRegion.Rollback()
                End If
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If mobjRegion.IsDirty Then
                mobjRegion.Save()
            End If
            Close()
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub btnRules_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnRules.Click
        Dim frmRules As New frmBusinessRules(mobjRegion.GetBusinessRules)
        frmRules.ShowDialog()
        frmRules = Nothing
    End Sub

End Class
