Option Strict On
Option Explicit On 

Public Class frmEditBase
    Inherits System.Windows.Forms.Form

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
    Protected WithEvents btnOK As System.Windows.Forms.Button
    Protected WithEvents btnCancel As System.Windows.Forms.Button
    Protected WithEvents btnRules As System.Windows.Forms.Button
    Protected WithEvents erpMain As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEditBase))
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnRules = New System.Windows.Forms.Button
        Me.erpMain = New System.Windows.Forms.ErrorProvider
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(264, 96)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "&OK"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(352, 96)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "&Cancel"
        '
        'btnRules
        '
        Me.btnRules.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRules.Image = CType(resources.GetObject("btnRules.Image"), System.Drawing.Image)
        Me.btnRules.Location = New System.Drawing.Point(8, 96)
        Me.btnRules.Name = "btnRules"
        Me.btnRules.Size = New System.Drawing.Size(24, 23)
        Me.btnRules.TabIndex = 2
        '
        'erpMain
        '
        Me.erpMain.ContainerControl = Me
        '
        'frmEditBase
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(432, 125)
        Me.Controls.Add(Me.btnRules)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEditBase"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmEditBase"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Close()
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub


#Region " Error Logger"
    Protected Sub LogException(ByVal exc As Exception)
        Dim objLogErr As New NorthwindUC.ErrorLogging.LogError

        Try
            objLogErr.LogException(exc)
            MessageBox.Show("The NorthwindTrader application generated " _
            & "the following error:" & ControlChars.CrLf & exc.Message, _
            "NorthwindTrader Error", MessageBoxButtons.OK, _
            MessageBoxIcon.Error)
        Catch excNew As Exception
            Dim objErrorEvent As NorthwindUC.ErrorLogging.LogErrorEvent
            objErrorEvent = objErrorEvent.getInstance
            objErrorEvent.LogErr(exc)
            objErrorEvent.LogErr(excNew)
            objErrorEvent = Nothing
            MessageBox.Show("The NorthwindTrader application generated " _
            & "the following critical error: " & excNew.Message, _
            "NorthwindTrader Error", MessageBoxButtons.OK, _
            MessageBoxIcon.Error)
        Finally
            objLogErr = Nothing
        End Try
    End Sub
#End Region

End Class
