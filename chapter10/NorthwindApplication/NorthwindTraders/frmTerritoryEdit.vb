Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindUC

Public Class frmTerritoryEdit
    Inherits NorthwindTraders.frmEditBase

    Private WithEvents mobjTerritory As Territory

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef objTerritory As Territory)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mobjTerritory = objTerritory

        If Not mobjTerritory.IsValid Then
            btnok.Enabled = False
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
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents lblTerritory As System.Windows.Forms.Label
    Friend WithEvents lblRegion As System.Windows.Forms.Label
    Friend WithEvents txtTerritoryID As System.Windows.Forms.TextBox
    Friend WithEvents txtTerritory As System.Windows.Forms.TextBox
    Friend WithEvents cboRegion As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblID = New System.Windows.Forms.Label
        Me.lblTerritory = New System.Windows.Forms.Label
        Me.lblRegion = New System.Windows.Forms.Label
        Me.txtTerritoryID = New System.Windows.Forms.TextBox
        Me.txtTerritory = New System.Windows.Forms.TextBox
        Me.cboRegion = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(138, 106)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 6
        '
        'btnRules
        '
        Me.btnRules.Location = New System.Drawing.Point(8, 106)
        Me.btnRules.Name = "btnRules"
        Me.btnRules.TabIndex = 8
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(226, 106)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 7
        '
        'lblID
        '
        Me.lblID.Location = New System.Drawing.Point(16, 16)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(64, 16)
        Me.lblID.TabIndex = 0
        Me.lblID.Text = "Territory ID:"
        Me.lblID.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblTerritory
        '
        Me.lblTerritory.Location = New System.Drawing.Point(16, 48)
        Me.lblTerritory.Name = "lblTerritory"
        Me.lblTerritory.Size = New System.Drawing.Size(64, 16)
        Me.lblTerritory.TabIndex = 2
        Me.lblTerritory.Text = "Territory:"
        Me.lblTerritory.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblRegion
        '
        Me.lblRegion.Location = New System.Drawing.Point(16, 80)
        Me.lblRegion.Name = "lblRegion"
        Me.lblRegion.Size = New System.Drawing.Size(64, 16)
        Me.lblRegion.TabIndex = 4
        Me.lblRegion.Text = "Region:"
        Me.lblRegion.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'txtTerritoryID
        '
        Me.erpMain.SetIconAlignment(Me.txtTerritoryID, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtTerritoryID.Location = New System.Drawing.Point(104, 8)
        Me.txtTerritoryID.Name = "txtTerritoryID"
        Me.txtTerritoryID.TabIndex = 1
        Me.txtTerritoryID.Text = ""
        '
        'txtTerritory
        '
        Me.erpMain.SetIconAlignment(Me.txtTerritory, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtTerritory.Location = New System.Drawing.Point(104, 40)
        Me.txtTerritory.Name = "txtTerritory"
        Me.txtTerritory.Size = New System.Drawing.Size(192, 20)
        Me.txtTerritory.TabIndex = 3
        Me.txtTerritory.Text = ""
        '
        'cboRegion
        '
        Me.erpMain.SetIconAlignment(Me.cboRegion, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.cboRegion.Location = New System.Drawing.Point(104, 72)
        Me.cboRegion.Name = "cboRegion"
        Me.cboRegion.Size = New System.Drawing.Size(152, 21)
        Me.cboRegion.TabIndex = 5
        '
        'frmTerritoryEdit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(306, 135)
        Me.Controls.Add(Me.cboRegion)
        Me.Controls.Add(Me.txtTerritory)
        Me.Controls.Add(Me.txtTerritoryID)
        Me.Controls.Add(Me.lblRegion)
        Me.Controls.Add(Me.lblTerritory)
        Me.Controls.Add(Me.lblID)
        Me.Name = "frmTerritoryEdit"
        Me.Text = "Territory [Detail]"
        Me.Controls.SetChildIndex(Me.lblID, 0)
        Me.Controls.SetChildIndex(Me.lblTerritory, 0)
        Me.Controls.SetChildIndex(Me.lblRegion, 0)
        Me.Controls.SetChildIndex(Me.txtTerritoryID, 0)
        Me.Controls.SetChildIndex(Me.txtTerritory, 0)
        Me.Controls.SetChildIndex(Me.cboRegion, 0)
        Me.Controls.SetChildIndex(Me.btnOK, 0)
        Me.Controls.SetChildIndex(Me.btnCancel, 0)
        Me.Controls.SetChildIndex(Me.btnRules, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmTerritoryEdit_Load(ByVal sender As Object, _
ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DictEnt As DictionaryEntry
        Dim objRegion As Region
        Dim objRegionMgr As RegionMgr

        Try
            objRegionMgr = objRegionMgr.GetInstance
            For Each DictEnt In objRegionMgr
                objRegion = CType(DictEnt.Value, Region)
                cboRegion.Items.Add(objRegion)
            Next

            If mobjTerritory.TerritoryID <> "" Then
                txtTerritoryID.Text = mobjTerritory.TerritoryID
                txtTerritory.Text = mobjTerritory.TerritoryDescription
                cboRegion.SelectedItem = mobjTerritory.Region
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

#Region " Validate Events"

    Private Sub txtTerritoryID_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtTerritoryID.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjTerritory.TerritoryID = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtTerritory_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtTerritory.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjTerritory.TerritoryDescription = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub cboRegion_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRegion.Validated
        Dim cbo As ComboBox = CType(sender, ComboBox)

        Try
            mobjTerritory.Region = CType(cbo.SelectedItem, Region)
            erpmain.SetError(cbo, "")
        Catch exc As Exception
            erpmain.SetError(cbo, "")
            erpmain.SetError(cbo, exc.Message)
        End Try
    End Sub

#End Region

    Private Sub btnOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If mobjTerritory.IsDirty Then
                mobjTerritory.Save()
            End If
            Close()
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub mobjRegion_BrokenRule(ByVal IsBroken As Boolean) _
    Handles mobjTerritory.BrokenRule
        If IsBroken Then
            btnOK.Enabled = False
        Else
            btnOK.Enabled = True
        End If
    End Sub

    Private Sub btnRules_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnRules.Click
        Dim frmRules As New frmBusinessRules(mobjTerritory.GetBusinessRules)
        frmRules.ShowDialog()
        frmRules = Nothing
    End Sub

End Class
