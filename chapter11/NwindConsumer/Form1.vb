Option Explicit On 
Option Strict On

Public Class frmMain
    Inherits System.Windows.Forms.Form

    Private mds As DataSet

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents grdEmployees As System.Windows.Forms.DataGrid
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grdEmployees = New System.Windows.Forms.DataGrid
        Me.btnLoad = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.grdEmployees, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdEmployees
        '
        Me.grdEmployees.DataMember = ""
        Me.grdEmployees.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdEmployees.Location = New System.Drawing.Point(8, 8)
        Me.grdEmployees.Name = "grdEmployees"
        Me.grdEmployees.Size = New System.Drawing.Size(376, 224)
        Me.grdEmployees.TabIndex = 0
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(8, 240)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(96, 23)
        Me.btnLoad.TabIndex = 1
        Me.btnLoad.Text = "Load Employees"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(152, 240)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(88, 23)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "Edit Employee"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(304, 240)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(392, 266)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.grdEmployees)
        Me.Name = "frmMain"
        Me.Text = "Employee Web Service Consumer"
        CType(Me.grdEmployees, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnLoad_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnLoad.Click
        Dim objWS As localhost.Service1

        Try
            Cursor = Cursors.WaitCursor
            objWS = New localhost.Service1
            objWS.Credentials = System.Net.CredentialCache.DefaultCredentials
            mds = objWS.GetAllEmployees
            grdEmployees.DataSource = mds.Tables(0)
        Catch exc As Exception
            MessageBox.Show(exc.Message)
        Finally
            objWS = Nothing
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, _
ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim objWS As localhost.Service1
        Dim frm As frmDetails
        Dim sEmployee As localhost.structEmployee

        Try
            Cursor = Cursors.WaitCursor
            objWS = New localhost.Service1
            objWS.Credentials = System.Net.CredentialCache.DefaultCredentials
            With grdEmployees
                sEmployee = _
           objWS.GetEmployeeDetails(Convert.ToInt32(.Item(.CurrentRowIndex, _
           0)))
            End With
            objWS = Nothing
            frm = New frmDetails(sEmployee)
            frm.ShowDialog()
        Catch exc As Exception
            MessageBox.Show(exc.Message)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

End Class
