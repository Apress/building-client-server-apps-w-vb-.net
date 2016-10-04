Option Explicit On 
Option Strict On

Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Http

Public Class frmMain
    Inherits System.Windows.Forms.Form

    Private chan As HttpChannel

    Private mfrmRegionList As frmRegionList

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim props As New Collections.Specialized.ListDictionary
        props.Add("useDefaultCredentials", True)
        chan = New HttpChannel(props, New BinaryClientFormatterSinkProvider, New _
        BinaryServerFormatterSinkProvider)
        ChannelServices.RegisterChannel(chan)
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Test"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Report Errors"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.IsMdiContainer = True
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMain"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

#End Region

    Private Sub frmMain_Closed(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            ChannelServices.UnregisterChannel(chan)
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MenuItem1.Click
        Try
            Cursor = Cursors.WaitCursor
            mfrmRegionList = New frmRegionList
            mfrmRegionList.MdiParent = Me
            mfrmRegionList.Show()
        Catch exc As Exception
            LogException(exc)
        Finally
            Cursor = Cursors.Default
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

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim objReportErrors As frmReportErrors

        Try
            objReportErrors = New frmReportErrors
            objReportErrors.ShowDialog()
        Catch exc As Exception
            LogException(exc)
        Finally
            objReportErrors = Nothing
        End Try

    End Sub
End Class
