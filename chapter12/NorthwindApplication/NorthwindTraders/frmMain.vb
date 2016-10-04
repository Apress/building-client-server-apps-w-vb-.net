Option Explicit On 
Option Strict On

Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Http
Imports System.IO
Imports System.Reflection
Imports NorthwindTraders.NorthwindUC

Public Class frmMain
    Inherits System.Windows.Forms.Form

    Private chan As HttpChannel
    Private mdsMenuLoad As DataSet
    Private mfrmRegionList As frmRegionList
    Private WithEvents mobjEventLog As ErrorLogging.LogErrorEvent
    Private WithEvents mfrmTerritoryList As frmTerritoryList
    Private WithEvents mfrmEmployeeList As frmEmployeeList

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
        LoadMenus()
        mobjEventLog = mobjEventLog.getInstance
        pnlDate.Text = Now.ToShortDateString
        pnlUser.Text = Security.Principal.WindowsIdentity.GetCurrent.Name
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
    Friend WithEvents tlbMain As System.Windows.Forms.ToolBar
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents tlbCut As System.Windows.Forms.ToolBarButton
    Friend WithEvents tlbCopy As System.Windows.Forms.ToolBarButton
    Friend WithEvents tlbPaste As System.Windows.Forms.ToolBarButton
    Friend WithEvents tlbFind As System.Windows.Forms.ToolBarButton
    Friend WithEvents sbrMain As System.Windows.Forms.StatusBar
    Friend WithEvents pnlStatus As System.Windows.Forms.StatusBarPanel
    Friend WithEvents pnlUser As System.Windows.Forms.StatusBarPanel
    Friend WithEvents pnlErrors As System.Windows.Forms.StatusBarPanel
    Friend WithEvents pnlCapsLock As System.Windows.Forms.StatusBarPanel
    Friend WithEvents pnlNumLock As System.Windows.Forms.StatusBarPanel
    Friend WithEvents pnlDate As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.tlbMain = New System.Windows.Forms.ToolBar
        Me.tlbCut = New System.Windows.Forms.ToolBarButton
        Me.tlbCopy = New System.Windows.Forms.ToolBarButton
        Me.tlbPaste = New System.Windows.Forms.ToolBarButton
        Me.tlbFind = New System.Windows.Forms.ToolBarButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.sbrMain = New System.Windows.Forms.StatusBar
        Me.pnlStatus = New System.Windows.Forms.StatusBarPanel
        Me.pnlUser = New System.Windows.Forms.StatusBarPanel
        Me.pnlErrors = New System.Windows.Forms.StatusBarPanel
        Me.pnlCapsLock = New System.Windows.Forms.StatusBarPanel
        Me.pnlNumLock = New System.Windows.Forms.StatusBarPanel
        Me.pnlDate = New System.Windows.Forms.StatusBarPanel
        CType(Me.pnlStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlErrors, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlCapsLock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlNumLock, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tlbMain
        '
        Me.tlbMain.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tlbCut, Me.tlbCopy, Me.tlbPaste, Me.tlbFind})
        Me.tlbMain.DropDownArrows = True
        Me.tlbMain.ImageList = Me.ImageList1
        Me.tlbMain.Location = New System.Drawing.Point(0, 0)
        Me.tlbMain.Name = "tlbMain"
        Me.tlbMain.ShowToolTips = True
        Me.tlbMain.Size = New System.Drawing.Size(640, 28)
        Me.tlbMain.TabIndex = 1
        '
        'tlbCut
        '
        Me.tlbCut.ImageIndex = 0
        '
        'tlbCopy
        '
        Me.tlbCopy.ImageIndex = 1
        '
        'tlbPaste
        '
        Me.tlbPaste.ImageIndex = 2
        '
        'tlbFind
        '
        Me.tlbFind.ImageIndex = 3
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'sbrMain
        '
        Me.sbrMain.Location = New System.Drawing.Point(0, 279)
        Me.sbrMain.Name = "sbrMain"
        Me.sbrMain.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.pnlStatus, Me.pnlUser, Me.pnlErrors, Me.pnlCapsLock, Me.pnlNumLock, Me.pnlDate})
        Me.sbrMain.ShowPanels = True
        Me.sbrMain.Size = New System.Drawing.Size(640, 22)
        Me.sbrMain.TabIndex = 3
        Me.sbrMain.Text = "StatusBar1"
        '
        'pnlStatus
        '
        Me.pnlStatus.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.pnlStatus.MinWidth = 100
        Me.pnlStatus.Width = 494
        '
        'pnlUser
        '
        Me.pnlUser.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.pnlUser.Width = 10
        '
        'pnlErrors
        '
        Me.pnlErrors.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.pnlErrors.Width = 10
        '
        'pnlCapsLock
        '
        Me.pnlCapsLock.MinWidth = 50
        Me.pnlCapsLock.Style = System.Windows.Forms.StatusBarPanelStyle.OwnerDraw
        Me.pnlCapsLock.Width = 50
        '
        'pnlNumLock
        '
        Me.pnlNumLock.MinWidth = 50
        Me.pnlNumLock.Style = System.Windows.Forms.StatusBarPanelStyle.OwnerDraw
        Me.pnlNumLock.Width = 50
        '
        'pnlDate
        '
        Me.pnlDate.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.pnlDate.Width = 10
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(640, 301)
        Me.Controls.Add(Me.sbrMain)
        Me.Controls.Add(Me.tlbMain)
        Me.IsMdiContainer = True
        Me.KeyPreview = True
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMain"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.pnlStatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlErrors, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlCapsLock, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlNumLock, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmMain_Load(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        If mobjEventLog.ErrorCount > 0 Then
            ShowErrorIcon()
        End If
    End Sub

    Private Sub frmMain_Closed(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            ChannelServices.UnregisterChannel(chan)
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub LoadRegion()
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

    Private Sub LoadTerritories()
        Try
            If mfrmTerritoryList Is Nothing Then
                Cursor = Cursors.WaitCursor
                mfrmTerritoryList = New frmTerritoryList
                mfrmTerritoryList.MdiParent = Me
                mfrmTerritoryList.Show()
            Else
                mfrmTerritoryList.Focus()
            End If
        Catch exc As Exception
            LogException(exc)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LoadEmployees()
        Try
            If mfrmEmployeeList Is Nothing Then
                Cursor = Cursors.WaitCursor
                mfrmEmployeeList = New frmEmployeeList
                mfrmEmployeeList.MdiParent = Me
                mfrmEmployeeList.Show()
            Else
                mfrmEmployeeList.Focus()
            End If
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

    Private Sub mobjEventLog_ErrorLogged() Handles mobjEventLog.ErrorLogged
        ShowErrorIcon()
    End Sub

    Private Sub mobjEventLog_ErrorsCleared() Handles mobjEventLog.ErrorsCleared
        pnlErrors.Icon = Nothing
    End Sub

    Private Sub LoadErrors()
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

    Private Sub MainMenu_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs)
        Dim mnu As MenuItem = CType(sender, MenuItem)

        Select Case mnu.Text
            Case "&Regions"
                LoadRegion()
            Case "&Report Errors"
                LoadErrors()
            Case "Cu&t", "&Copy", "&Paste", "Select Al&l"
                CutCopyPaste(mnu.Text)
            Case "&Find..."
                Find(1)
            Case "Find &Next"
                Find(0)
            Case "&Territories"
                LoadTerritories()
            Case "&Employees"
                LoadEmployees()
            Case "E&xit"
                Close()
        End Select
    End Sub

    Private Sub CutCopyPaste(ByVal strMenu As String)
        Dim objCCPU As ICutCopyPaste
        Dim t As System.Type

        If ActiveMdiChild Is Nothing Then Exit Sub
        Try
            t = ActiveMdiChild.GetType
            If Not t.GetInterface("ICutCopyPaste") Is Nothing Then
                objCCPU = CType(Me.ActiveMdiChild, ICutCopyPaste)

                Select Case strMenu
                    Case "Cu&t"
                        objCCPU.Cut()
                    Case "&Copy"
                        objCCPU.Copy()
                    Case "&Paste"
                        objCCPU.Paste()
                    Case "Select Al&l"
                        objCCPU.SelectAll()
                End Select
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub LoadMenus()
        Try
            Dim objMenu As UIMenu
            objMenu = objMenu.getInstance

            mdsMenuLoad = objMenu.GetMenuStructure

            Dim dv As New DataView(mdsMenuLoad.Tables(0))
            Dim drv As DataRowView

            dv.RowFilter = "menu_under_id is null"
            dv.RowStateFilter = DataViewRowState.CurrentRows

            For Each drv In dv
                Dim mnuHeader As New _
                MenuItem(Convert.ToString(drv.Item("menu_caption")))
                AddItems(mnuHeader, Convert.ToInt32(drv.Item("menu_id")))
                MainMenu1.MenuItems.Add(mnuHeader)
            Next
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub AddItems(ByVal mnuTop As MenuItem, ByVal intMenuID As Integer)
        Dim dr As DataRowView
        Dim dv As New DataView(mdsMenuLoad.Tables(0))

        Try
            dv.RowFilter = "menu_under_id = " & intMenuID
            dv.RowStateFilter = DataViewRowState.CurrentRows
            If dv.Count > 0 Then
                For Each dr In dv
                    Dim mnuItem As New _
                    MenuItem(Convert.ToString(dr.Item("menu_caption")))
                    If Not IsDBNull(dr.Item("menu_shortcut")) Then
                        mnuItem.Shortcut = _
                        CType([Enum].Parse(GetType(Shortcut), _
                        Convert.ToString(dr.Item("menu_shortcut"))), _
                        Shortcut)
                    End If

                    If Convert.ToInt32(dr.Item("enabled")) = 0 Then
                        mnuItem.Enabled = False
                    Else
                        mnuItem.Enabled = True
                    End If

                    If Convert.ToInt32(dr.Item("checked")) = 0 Then
                        mnuItem.Checked = False
                    Else
                        mnuItem.Checked = True
                    End If

                    AddItems(mnuItem, Convert.ToInt32(dr.Item("menu_id")))
                    AddHandler mnuItem.Click, AddressOf MainMenu_Click
                    mnuTop.MenuItems.Add(mnuItem)
                Next
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub Find(ByVal intType As Integer)
        Dim objFind As IFind
        Dim t As System.Type

        If ActiveMdiChild Is Nothing Then Exit Sub
        Try
            t = ActiveMdiChild.GetType
            If Not t.GetInterface("IFind") Is Nothing Then
                objFind = CType(Me.ActiveMdiChild, IFind)
                If intType = 1 Then
                    objFind.Find()
                Else
                    objFind.FindNext()
                End If
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub tlbMain_ButtonClick(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) _
    Handles tlbMain.ButtonClick

        Select Case tlbMain.Buttons.IndexOf(e.Button)
            Case 0
                CutCopyPaste("Cu&t")
            Case 1
                CutCopyPaste("&Copy")
            Case 2
                CutCopyPaste("&Paste")
            Case 3
                Find(1)
        End Select
    End Sub

    Private Sub sbrMain_DrawItem(ByVal sender As System.Object, _
    ByVal sbdevent As System.Windows.Forms.StatusBarDrawItemEventArgs) _
    Handles sbrMain.DrawItem

        Dim intValue(0) As Integer
        Dim ba As BitArray
        Dim fnt As New Font(New FontFamily("Arial"), 10, FontStyle.Regular)
        Dim strText As String
        Dim clr As Color

        If sbdevent.Panel Is pnlCapsLock Then
            intValue(0) = Win32APIcalls.GetKeyState(Keys.CapsLock)
            ba = New BitArray(intValue)
            strText = "CAPS"
            If ba.Item(0) = False Then
                clr = Color.DarkGray
            Else
                clr = Color.Black
            End If
        End If

        If sbdevent.Panel Is pnlNumLock Then
            intValue(0) = Win32APIcalls.GetKeyState(Keys.NumLock)
            ba = New BitArray(intValue)
            strText = "NUM"
            If ba.Item(0) = False Then
                clr = Color.DarkGray
            Else
                clr = Color.Black
            End If
        End If

        sbdevent.Graphics.DrawString(strText, fnt, New SolidBrush(clr), _
        sbdevent.Graphics.VisibleClipBounds.Location.X, _
        sbdevent.Graphics.VisibleClipBounds.Location.Y)
    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.CapsLock Or e.KeyCode = Keys.NumLock Then
            sbrMain.Refresh()
        End If
    End Sub

    Private Sub ShowErrorIcon()
        Dim asm As [Assembly] = [Assembly].GetExecutingAssembly
        Dim iStream As Stream

        iStream = asm.GetManifestResourceStream _
        ("NorthwindTraders.W95MBX01.ICO")
        pnlErrors.Icon = New Icon(iStream)
    End Sub

    Private Sub mfrmTerritoryList_Closed(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles mfrmTerritoryList.Closed
        mfrmTerritoryList = Nothing
    End Sub

    Private Sub mfrmEmployeeList_Closed(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles mfrmEmployeeList.Closed
        mfrmEmployeeList = Nothing
    End Sub
End Class
