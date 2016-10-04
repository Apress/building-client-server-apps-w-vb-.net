Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindUC.ErrorLogging

Public Class frmReportErrors
    Inherits System.Windows.Forms.Form

    Private mobjLogErr As LogErrorEvent
    Private mobjErrArray() As structLoggedError

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lstErrors As System.Windows.Forms.ListBox
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lstErrors = New System.Windows.Forms.ListBox
        Me.btnSend = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.White
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(-8, -8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(304, 64)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(40, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Send the error log to tech support for diagnosis"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(264, 32)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "The following errors will be sent to the technical support personnel:"
        '
        'lstErrors
        '
        Me.lstErrors.Location = New System.Drawing.Point(8, 96)
        Me.lstErrors.Name = "lstErrors"
        Me.lstErrors.Size = New System.Drawing.Size(264, 147)
        Me.lstErrors.TabIndex = 2
        '
        'btnSend
        '
        Me.btnSend.Location = New System.Drawing.Point(120, 256)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.TabIndex = 3
        Me.btnSend.Text = "&Send"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(200, 256)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(168, 16)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Report Errors Wizard"
        '
        'frmReportErrors
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(280, 287)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.lstErrors)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReportErrors"
        Me.Text = "Northwind Wizard"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmReportErrors_Load(ByVal sender As Object, _
ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer

        Try
            mobjLogErr = mobjLogErr.getInstance
            mobjErrArray = mobjLogErr.RetrieveErrors

            For i = 0 To mobjErrArray.Length - 1
                lstErrors.Items.Add(mobjErrArray(i).Message)
            Next

            If lstErrors.Items.Count = 0 Then
                btnSend.Enabled = False
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub btnSend_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnSend.Click
        Try
            Cursor = Cursors.WaitCursor
            mobjLogErr.SendErrors(mobjErrArray)
        Catch exc As Exception
            LogException(exc)
        Finally
            Cursor = Cursors.Default
            Me.Close()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Close()
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub frmReportErrors_Closed(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            mobjLogErr = Nothing
            mobjErrArray = Nothing
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
