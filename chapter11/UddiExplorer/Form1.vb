Option Explicit On 
Option Strict On

Imports Microsoft
Imports Microsoft.Uddi
Imports Microsoft.Uddi.Api

Public Class Form1
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
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tvwResults As System.Windows.Forms.TreeView
    Friend WithEvents txtDetails As System.Windows.Forms.TextBox
    Friend WithEvents btnDetails As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tvwResults = New System.Windows.Forms.TreeView
        Me.btnSearch = New System.Windows.Forms.Button
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtDetails = New System.Windows.Forms.TextBox
        Me.btnDetails = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'tvwResults
        '
        Me.tvwResults.ImageIndex = -1
        Me.tvwResults.Location = New System.Drawing.Point(16, 80)
        Me.tvwResults.Name = "tvwResults"
        Me.tvwResults.SelectedImageIndex = -1
        Me.tvwResults.Size = New System.Drawing.Size(208, 224)
        Me.tvwResults.TabIndex = 0
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(16, 312)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.TabIndex = 1
        Me.btnSearch.Text = "&Search"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(16, 32)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(208, 20)
        Me.txtName.TabIndex = 2
        Me.txtName.Text = "NorthwindTraders"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Business"
        '
        'txtDetails
        '
        Me.txtDetails.Location = New System.Drawing.Point(256, 80)
        Me.txtDetails.Multiline = True
        Me.txtDetails.Name = "txtDetails"
        Me.txtDetails.Size = New System.Drawing.Size(240, 224)
        Me.txtDetails.TabIndex = 4
        Me.txtDetails.Text = ""
        '
        'btnDetails
        '
        Me.btnDetails.Location = New System.Drawing.Point(152, 312)
        Me.btnDetails.Name = "btnDetails"
        Me.btnDetails.TabIndex = 5
        Me.btnDetails.Text = "&Details"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(424, 312)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "&Close"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Services"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(256, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(184, 16)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Service Description"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(512, 342)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnDetails)
        Me.Controls.Add(Me.txtDetails)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.tvwResults)
        Me.Name = "Form1"
        Me.Text = "UDDI Explorer"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Try
            Cursor = Cursors.WaitCursor
            Inquire.Url = "http://localhost/uddipublic/inquire.asmx"

            Dim fb As New FindBusiness
            Dim bList As BusinessList
            Dim bus As Business.BusinessInfo

            fb.Names.Add(Me.txtName.Text)
            bList = fb.Send

            For Each bus In bList.BusinessInfos
                Dim ndeBus As New TreeNode(bus.Name)
                ndeBus.Tag = bus.BusinessKey

                Dim gBusinessDetail As New GetBusinessDetail
                Dim bDetail As BusinessDetail

                gBusinessDetail.BusinessKeys.Add(bus.BusinessKey)
                bDetail = gBusinessDetail.Send

                Dim i As Integer
                Dim gServices As New GetServiceDetail
                Dim bService As ServiceDetail
                Dim bEntity As Business.BusinessEntity

                For Each bEntity In bDetail.BusinessEntities
                    For i = 0 To bEntity.BusinessServices.Count - 1
                        Dim nde As New TreeNode(bEntity.BusinessServices(i).Names(0).Text)
                        nde.Tag = bEntity.BusinessServices(i).ServiceKey
                        ndeBus.Nodes.Add(nde)
                    Next
                Next

                tvwResults.Nodes.Add(ndeBus)
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnDetails_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnDetails.Click
        Dim gServiceDetail As New GetServiceDetail
        Dim bService As ServiceDetail
        Dim i, j As Integer

        gServiceDetail.ServiceKeys.Add(tvwResults.SelectedNode.Tag.ToString)
        bService = gServiceDetail.Send

        txtDetails.Text = ""

        For i = 0 To bService.BusinessServices.Count - 1
            txtDetails.Text = "Service Name: " & _
            bService.BusinessServices(i).Names(0).Text
            txtDetails.Text += ControlChars.CrLf

            txtDetails.Text += "Descriptions: " & ControlChars.CrLf
            For j = 0 To bService.BusinessServices(i).Descriptions.Count - 1
                txtDetails.Text += bService.BusinessServices(i).Descriptions(j).Text
                txtDetails.Text += ControlChars.CrLf
            Next

            txtDetails.Text += "Category Bag: " & ControlChars.CrLf
            For j = 0 To bService.BusinessServices(i).CategoryBag.Count - 1
                txtDetails.Text += "Key Name: " & _
                bService.BusinessServices(i).CategoryBag(j).KeyName & ControlChars.CrLf
                txtDetails.Text += "Key Value: " & _
                bService.BusinessServices(i).CategoryBag(j).KeyValue & ControlChars.CrLf
                txtDetails.Text += "tModel Key: " & _
                bService.BusinessServices(i).CategoryBag(j).TModelKey & ControlChars.CrLf
                txtDetails.Text += ControlChars.CrLf
            Next

            txtDetails.Text += "Business Key: " & _
            bService.BusinessServices(i).BusinessKey
            txtDetails.Text += ControlChars.CrLf
            txtDetails.Text += "Service Key: " & _
            bService.BusinessServices(i).ServiceKey
        Next
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub
End Class
