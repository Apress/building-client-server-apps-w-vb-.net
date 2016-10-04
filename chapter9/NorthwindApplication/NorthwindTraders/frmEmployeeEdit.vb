Option Strict On
Option Explicit On 

Imports NorthwindTraders.NorthwindUC
Imports System.Globalization
Imports System.Resources
Imports System.Threading

Public Class frmEmployeeEdit
    Inherits frmEditBase

    Private WithEvents mobjEmployee As Employee
    Private mblnErrors As Boolean = False
    Private mobjTerritoryMgr As TerritoryMgr

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef objEmployee As Employee, ByRef objTerritoryMgr As TerritoryMgr)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mobjEmployee = objEmployee
        mobjTerritoryMgr = objTerritoryMgr

        If Not mobjEmployee.IsValid Then
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
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents txtCountry As System.Windows.Forms.TextBox
    Friend WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Friend WithEvents cboCourtesy As System.Windows.Forms.ComboBox
    Friend WithEvents cboReportsTo As System.Windows.Forms.ComboBox
    Friend WithEvents dtpBirthDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpHireDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lstTerritories As System.Windows.Forms.ListBox
    Friend WithEvents lstAvailable As System.Windows.Forms.ListBox
    Friend WithEvents txtRegion As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents picPhoto As System.Windows.Forms.PictureBox
    Friend WithEvents txtExtension As System.Windows.Forms.TextBox
    Friend WithEvents txtHomePhone As System.Windows.Forms.TextBox
    Friend WithEvents btnPhoto As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents rtbNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblCourtesy As System.Windows.Forms.Label
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents lblReportsTo As System.Windows.Forms.Label
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblHireDate As System.Windows.Forms.Label
    Friend WithEvents lblBirthDate As System.Windows.Forms.Label
    Friend WithEvents lblExtension As System.Windows.Forms.Label
    Friend WithEvents lblHomePhone As System.Windows.Forms.Label
    Friend WithEvents lblCountry As System.Windows.Forms.Label
    Friend WithEvents lblPostalCode As System.Windows.Forms.Label
    Friend WithEvents lblRegion As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblAvailable As System.Windows.Forms.Label
    Friend WithEvents lblAssigned As System.Windows.Forms.Label
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents lblPhoto As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEmployeeEdit))
        Me.txtTitle = New System.Windows.Forms.TextBox
        Me.txtLastName = New System.Windows.Forms.TextBox
        Me.txtFirstName = New System.Windows.Forms.TextBox
        Me.txtCountry = New System.Windows.Forms.TextBox
        Me.txtPostalCode = New System.Windows.Forms.TextBox
        Me.cboCourtesy = New System.Windows.Forms.ComboBox
        Me.cboReportsTo = New System.Windows.Forms.ComboBox
        Me.dtpBirthDate = New System.Windows.Forms.DateTimePicker
        Me.dtpHireDate = New System.Windows.Forms.DateTimePicker
        Me.lstTerritories = New System.Windows.Forms.ListBox
        Me.lstAvailable = New System.Windows.Forms.ListBox
        Me.txtRegion = New System.Windows.Forms.TextBox
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.txtCity = New System.Windows.Forms.TextBox
        Me.picPhoto = New System.Windows.Forms.PictureBox
        Me.txtExtension = New System.Windows.Forms.TextBox
        Me.txtHomePhone = New System.Windows.Forms.TextBox
        Me.lblCourtesy = New System.Windows.Forms.Label
        Me.lblLastName = New System.Windows.Forms.Label
        Me.lblFirstName = New System.Windows.Forms.Label
        Me.lblReportsTo = New System.Windows.Forms.Label
        Me.lblTitle = New System.Windows.Forms.Label
        Me.lblHireDate = New System.Windows.Forms.Label
        Me.lblBirthDate = New System.Windows.Forms.Label
        Me.lblExtension = New System.Windows.Forms.Label
        Me.lblHomePhone = New System.Windows.Forms.Label
        Me.lblCountry = New System.Windows.Forms.Label
        Me.lblPostalCode = New System.Windows.Forms.Label
        Me.lblRegion = New System.Windows.Forms.Label
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblAddress = New System.Windows.Forms.Label
        Me.lblAvailable = New System.Windows.Forms.Label
        Me.lblAssigned = New System.Windows.Forms.Label
        Me.lblNotes = New System.Windows.Forms.Label
        Me.lblPhoto = New System.Windows.Forms.Label
        Me.btnPhoto = New System.Windows.Forms.Button
        Me.btnRemove = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.rtbNotes = New System.Windows.Forms.RichTextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(610, 442)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 20
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(698, 442)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 21
        '
        'btnRules
        '
        Me.btnRules.Location = New System.Drawing.Point(8, 442)
        Me.btnRules.Name = "btnRules"
        Me.btnRules.TabIndex = 22
        '
        'txtTitle
        '
        Me.erpMain.SetIconAlignment(Me.txtTitle, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtTitle.Location = New System.Drawing.Point(16, 80)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(128, 20)
        Me.txtTitle.TabIndex = 3
        Me.txtTitle.Text = ""
        '
        'txtLastName
        '
        Me.erpMain.SetIconAlignment(Me.txtLastName, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtLastName.Location = New System.Drawing.Point(232, 24)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(128, 20)
        Me.txtLastName.TabIndex = 2
        Me.txtLastName.Text = ""
        '
        'txtFirstName
        '
        Me.erpMain.SetIconAlignment(Me.txtFirstName, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtFirstName.Location = New System.Drawing.Point(112, 24)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(88, 20)
        Me.txtFirstName.TabIndex = 1
        Me.txtFirstName.Text = ""
        '
        'txtCountry
        '
        Me.erpMain.SetIconAlignment(Me.txtCountry, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtCountry.Location = New System.Drawing.Point(112, 400)
        Me.txtCountry.Name = "txtCountry"
        Me.txtCountry.Size = New System.Drawing.Size(168, 20)
        Me.txtCountry.TabIndex = 13
        Me.txtCountry.Text = ""
        '
        'txtPostalCode
        '
        Me.erpMain.SetIconAlignment(Me.txtPostalCode, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtPostalCode.Location = New System.Drawing.Point(112, 368)
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.Size = New System.Drawing.Size(168, 20)
        Me.txtPostalCode.TabIndex = 12
        Me.txtPostalCode.Text = ""
        '
        'cboCourtesy
        '
        Me.cboCourtesy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.erpMain.SetIconAlignment(Me.cboCourtesy, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.cboCourtesy.Items.AddRange(New Object() {"Mr.", "Mrs.", "Dr.", "Ms."})
        Me.cboCourtesy.Location = New System.Drawing.Point(16, 24)
        Me.cboCourtesy.Name = "cboCourtesy"
        Me.cboCourtesy.Size = New System.Drawing.Size(72, 21)
        Me.cboCourtesy.TabIndex = 0
        '
        'cboReportsTo
        '
        Me.cboReportsTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.erpMain.SetIconAlignment(Me.cboReportsTo, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.cboReportsTo.Location = New System.Drawing.Point(168, 80)
        Me.cboReportsTo.Name = "cboReportsTo"
        Me.cboReportsTo.Size = New System.Drawing.Size(152, 21)
        Me.cboReportsTo.TabIndex = 4
        '
        'dtpBirthDate
        '
        Me.dtpBirthDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.erpMain.SetIconAlignment(Me.dtpBirthDate, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.dtpBirthDate.Location = New System.Drawing.Point(24, 136)
        Me.dtpBirthDate.Name = "dtpBirthDate"
        Me.dtpBirthDate.Size = New System.Drawing.Size(88, 20)
        Me.dtpBirthDate.TabIndex = 5
        '
        'dtpHireDate
        '
        Me.dtpHireDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.erpMain.SetIconAlignment(Me.dtpHireDate, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.dtpHireDate.Location = New System.Drawing.Point(136, 136)
        Me.dtpHireDate.Name = "dtpHireDate"
        Me.dtpHireDate.Size = New System.Drawing.Size(88, 20)
        Me.dtpHireDate.TabIndex = 6
        '
        'lstTerritories
        '
        Me.erpMain.SetIconAlignment(Me.lstTerritories, System.Windows.Forms.ErrorIconAlignment.TopLeft)
        Me.lstTerritories.Location = New System.Drawing.Point(416, 272)
        Me.lstTerritories.Name = "lstTerritories"
        Me.lstTerritories.Size = New System.Drawing.Size(120, 147)
        Me.lstTerritories.Sorted = True
        Me.lstTerritories.TabIndex = 16
        '
        'lstAvailable
        '
        Me.erpMain.SetIconAlignment(Me.lstAvailable, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lstAvailable.Location = New System.Drawing.Point(576, 272)
        Me.lstAvailable.Name = "lstAvailable"
        Me.lstAvailable.Size = New System.Drawing.Size(120, 147)
        Me.lstAvailable.Sorted = True
        Me.lstAvailable.TabIndex = 19
        '
        'txtRegion
        '
        Me.erpMain.SetIconAlignment(Me.txtRegion, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtRegion.Location = New System.Drawing.Point(112, 336)
        Me.txtRegion.Name = "txtRegion"
        Me.txtRegion.Size = New System.Drawing.Size(168, 20)
        Me.txtRegion.TabIndex = 11
        Me.txtRegion.Text = ""
        '
        'txtAddress
        '
        Me.erpMain.SetIconAlignment(Me.txtAddress, System.Windows.Forms.ErrorIconAlignment.TopLeft)
        Me.txtAddress.Location = New System.Drawing.Point(112, 256)
        Me.txtAddress.Multiline = True
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(248, 40)
        Me.txtAddress.TabIndex = 9
        Me.txtAddress.Text = ""
        '
        'txtCity
        '
        Me.erpMain.SetIconAlignment(Me.txtCity, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtCity.Location = New System.Drawing.Point(112, 304)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(168, 20)
        Me.txtCity.TabIndex = 10
        Me.txtCity.Text = ""
        '
        'picPhoto
        '
        Me.picPhoto.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.erpMain.SetIconAlignment(Me.picPhoto, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.picPhoto.Location = New System.Drawing.Point(384, 40)
        Me.picPhoto.Name = "picPhoto"
        Me.picPhoto.Size = New System.Drawing.Size(176, 184)
        Me.picPhoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picPhoto.TabIndex = 17
        Me.picPhoto.TabStop = False
        '
        'txtExtension
        '
        Me.erpMain.SetIconAlignment(Me.txtExtension, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtExtension.Location = New System.Drawing.Point(176, 192)
        Me.txtExtension.Name = "txtExtension"
        Me.txtExtension.Size = New System.Drawing.Size(64, 20)
        Me.txtExtension.TabIndex = 8
        Me.txtExtension.Text = ""
        '
        'txtHomePhone
        '
        Me.erpMain.SetIconAlignment(Me.txtHomePhone, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtHomePhone.Location = New System.Drawing.Point(24, 192)
        Me.txtHomePhone.Name = "txtHomePhone"
        Me.txtHomePhone.Size = New System.Drawing.Size(120, 20)
        Me.txtHomePhone.TabIndex = 7
        Me.txtHomePhone.Text = ""
        '
        'lblCourtesy
        '
        Me.erpMain.SetIconAlignment(Me.lblCourtesy, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblCourtesy.Location = New System.Drawing.Point(16, 48)
        Me.lblCourtesy.Name = "lblCourtesy"
        Me.lblCourtesy.Size = New System.Drawing.Size(72, 16)
        Me.lblCourtesy.TabIndex = 21
        Me.lblCourtesy.Text = "Courtesy"
        '
        'lblLastName
        '
        Me.erpMain.SetIconAlignment(Me.lblLastName, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblLastName.Location = New System.Drawing.Point(232, 48)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(128, 16)
        Me.lblLastName.TabIndex = 22
        Me.lblLastName.Text = "Last Name"
        '
        'lblFirstName
        '
        Me.erpMain.SetIconAlignment(Me.lblFirstName, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblFirstName.Location = New System.Drawing.Point(112, 48)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(88, 16)
        Me.lblFirstName.TabIndex = 23
        Me.lblFirstName.Text = "First Name"
        '
        'lblReportsTo
        '
        Me.erpMain.SetIconAlignment(Me.lblReportsTo, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblReportsTo.Location = New System.Drawing.Point(168, 104)
        Me.lblReportsTo.Name = "lblReportsTo"
        Me.lblReportsTo.Size = New System.Drawing.Size(152, 16)
        Me.lblReportsTo.TabIndex = 24
        Me.lblReportsTo.Text = "Reports To"
        '
        'lblTitle
        '
        Me.erpMain.SetIconAlignment(Me.lblTitle, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblTitle.Location = New System.Drawing.Point(16, 104)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(128, 16)
        Me.lblTitle.TabIndex = 25
        Me.lblTitle.Text = "Title"
        '
        'lblHireDate
        '
        Me.erpMain.SetIconAlignment(Me.lblHireDate, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblHireDate.Location = New System.Drawing.Point(136, 160)
        Me.lblHireDate.Name = "lblHireDate"
        Me.lblHireDate.Size = New System.Drawing.Size(72, 16)
        Me.lblHireDate.TabIndex = 26
        Me.lblHireDate.Text = "Hire Date"
        '
        'lblBirthDate
        '
        Me.erpMain.SetIconAlignment(Me.lblBirthDate, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblBirthDate.Location = New System.Drawing.Point(24, 160)
        Me.lblBirthDate.Name = "lblBirthDate"
        Me.lblBirthDate.Size = New System.Drawing.Size(72, 16)
        Me.lblBirthDate.TabIndex = 27
        Me.lblBirthDate.Text = "Birth Date"
        '
        'lblExtension
        '
        Me.erpMain.SetIconAlignment(Me.lblExtension, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblExtension.Location = New System.Drawing.Point(176, 216)
        Me.lblExtension.Name = "lblExtension"
        Me.lblExtension.Size = New System.Drawing.Size(64, 16)
        Me.lblExtension.TabIndex = 28
        Me.lblExtension.Text = "Extension"
        '
        'lblHomePhone
        '
        Me.erpMain.SetIconAlignment(Me.lblHomePhone, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblHomePhone.Location = New System.Drawing.Point(24, 216)
        Me.lblHomePhone.Name = "lblHomePhone"
        Me.lblHomePhone.Size = New System.Drawing.Size(72, 16)
        Me.lblHomePhone.TabIndex = 29
        Me.lblHomePhone.Text = "Home Phone"
        '
        'lblCountry
        '
        Me.erpMain.SetIconAlignment(Me.lblCountry, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblCountry.Location = New System.Drawing.Point(16, 408)
        Me.lblCountry.Name = "lblCountry"
        Me.lblCountry.Size = New System.Drawing.Size(72, 16)
        Me.lblCountry.TabIndex = 30
        Me.lblCountry.Text = "Country:"
        Me.lblCountry.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPostalCode
        '
        Me.erpMain.SetIconAlignment(Me.lblPostalCode, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblPostalCode.Location = New System.Drawing.Point(16, 376)
        Me.lblPostalCode.Name = "lblPostalCode"
        Me.lblPostalCode.Size = New System.Drawing.Size(72, 16)
        Me.lblPostalCode.TabIndex = 31
        Me.lblPostalCode.Text = "Postal Code:"
        Me.lblPostalCode.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblRegion
        '
        Me.erpMain.SetIconAlignment(Me.lblRegion, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblRegion.Location = New System.Drawing.Point(16, 344)
        Me.lblRegion.Name = "lblRegion"
        Me.lblRegion.Size = New System.Drawing.Size(72, 16)
        Me.lblRegion.TabIndex = 32
        Me.lblRegion.Text = "Region:"
        Me.lblRegion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCity
        '
        Me.erpMain.SetIconAlignment(Me.lblCity, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblCity.Location = New System.Drawing.Point(16, 312)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(72, 16)
        Me.lblCity.TabIndex = 33
        Me.lblCity.Text = "City:"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAddress
        '
        Me.erpMain.SetIconAlignment(Me.lblAddress, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblAddress.Location = New System.Drawing.Point(16, 256)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(72, 16)
        Me.lblAddress.TabIndex = 34
        Me.lblAddress.Text = "Address:"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAvailable
        '
        Me.erpMain.SetIconAlignment(Me.lblAvailable, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblAvailable.Location = New System.Drawing.Point(576, 256)
        Me.lblAvailable.Name = "lblAvailable"
        Me.lblAvailable.Size = New System.Drawing.Size(112, 16)
        Me.lblAvailable.TabIndex = 35
        Me.lblAvailable.Text = "Available Territories"
        '
        'lblAssigned
        '
        Me.erpMain.SetIconAlignment(Me.lblAssigned, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblAssigned.Location = New System.Drawing.Point(416, 256)
        Me.lblAssigned.Name = "lblAssigned"
        Me.lblAssigned.Size = New System.Drawing.Size(112, 16)
        Me.lblAssigned.TabIndex = 36
        Me.lblAssigned.Text = "Assigned Territories"
        '
        'lblNotes
        '
        Me.erpMain.SetIconAlignment(Me.lblNotes, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblNotes.Location = New System.Drawing.Point(584, 24)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(72, 16)
        Me.lblNotes.TabIndex = 37
        Me.lblNotes.Text = "Notes"
        '
        'lblPhoto
        '
        Me.erpMain.SetIconAlignment(Me.lblPhoto, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.lblPhoto.Location = New System.Drawing.Point(384, 24)
        Me.lblPhoto.Name = "lblPhoto"
        Me.lblPhoto.Size = New System.Drawing.Size(72, 16)
        Me.lblPhoto.TabIndex = 38
        Me.lblPhoto.Text = "Photo"
        '
        'btnPhoto
        '
        Me.erpMain.SetIconAlignment(Me.btnPhoto, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.btnPhoto.Image = CType(resources.GetObject("btnPhoto.Image"), System.Drawing.Image)
        Me.btnPhoto.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPhoto.Location = New System.Drawing.Point(536, 16)
        Me.btnPhoto.Name = "btnPhoto"
        Me.btnPhoto.Size = New System.Drawing.Size(24, 23)
        Me.btnPhoto.TabIndex = 14
        '
        'btnRemove
        '
        Me.erpMain.SetIconAlignment(Me.btnRemove, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.btnRemove.Location = New System.Drawing.Point(544, 352)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(24, 23)
        Me.btnRemove.TabIndex = 18
        Me.btnRemove.Text = ">"
        '
        'btnAdd
        '
        Me.erpMain.SetIconAlignment(Me.btnAdd, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.btnAdd.Location = New System.Drawing.Point(544, 312)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(24, 23)
        Me.btnAdd.TabIndex = 17
        Me.btnAdd.Text = "<"
        '
        'rtbNotes
        '
        Me.erpMain.SetIconAlignment(Me.rtbNotes, System.Windows.Forms.ErrorIconAlignment.TopLeft)
        Me.rtbNotes.Location = New System.Drawing.Point(584, 40)
        Me.rtbNotes.Name = "rtbNotes"
        Me.rtbNotes.Size = New System.Drawing.Size(184, 184)
        Me.rtbNotes.TabIndex = 15
        Me.rtbNotes.Text = ""
        '
        'frmEmployeeEdit
        '
        Me.AcceptButton = Nothing
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(778, 472)
        Me.Controls.Add(Me.rtbNotes)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.btnPhoto)
        Me.Controls.Add(Me.lblPhoto)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.lblAssigned)
        Me.Controls.Add(Me.lblAvailable)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblRegion)
        Me.Controls.Add(Me.lblPostalCode)
        Me.Controls.Add(Me.lblCountry)
        Me.Controls.Add(Me.lblHomePhone)
        Me.Controls.Add(Me.lblExtension)
        Me.Controls.Add(Me.lblBirthDate)
        Me.Controls.Add(Me.lblHireDate)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.lblReportsTo)
        Me.Controls.Add(Me.lblFirstName)
        Me.Controls.Add(Me.lblLastName)
        Me.Controls.Add(Me.lblCourtesy)
        Me.Controls.Add(Me.txtHomePhone)
        Me.Controls.Add(Me.txtExtension)
        Me.Controls.Add(Me.picPhoto)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.txtRegion)
        Me.Controls.Add(Me.lstAvailable)
        Me.Controls.Add(Me.lstTerritories)
        Me.Controls.Add(Me.dtpHireDate)
        Me.Controls.Add(Me.dtpBirthDate)
        Me.Controls.Add(Me.cboReportsTo)
        Me.Controls.Add(Me.cboCourtesy)
        Me.Controls.Add(Me.txtPostalCode)
        Me.Controls.Add(Me.txtCountry)
        Me.Controls.Add(Me.txtFirstName)
        Me.Controls.Add(Me.txtLastName)
        Me.Controls.Add(Me.txtTitle)
        Me.Name = "frmEmployeeEdit"
        Me.Text = "Employee Edit [Details]"
        Me.Controls.SetChildIndex(Me.txtTitle, 0)
        Me.Controls.SetChildIndex(Me.txtLastName, 0)
        Me.Controls.SetChildIndex(Me.txtFirstName, 0)
        Me.Controls.SetChildIndex(Me.txtCountry, 0)
        Me.Controls.SetChildIndex(Me.txtPostalCode, 0)
        Me.Controls.SetChildIndex(Me.cboCourtesy, 0)
        Me.Controls.SetChildIndex(Me.cboReportsTo, 0)
        Me.Controls.SetChildIndex(Me.dtpBirthDate, 0)
        Me.Controls.SetChildIndex(Me.dtpHireDate, 0)
        Me.Controls.SetChildIndex(Me.lstTerritories, 0)
        Me.Controls.SetChildIndex(Me.lstAvailable, 0)
        Me.Controls.SetChildIndex(Me.txtRegion, 0)
        Me.Controls.SetChildIndex(Me.txtAddress, 0)
        Me.Controls.SetChildIndex(Me.txtCity, 0)
        Me.Controls.SetChildIndex(Me.picPhoto, 0)
        Me.Controls.SetChildIndex(Me.txtExtension, 0)
        Me.Controls.SetChildIndex(Me.txtHomePhone, 0)
        Me.Controls.SetChildIndex(Me.lblCourtesy, 0)
        Me.Controls.SetChildIndex(Me.lblLastName, 0)
        Me.Controls.SetChildIndex(Me.lblFirstName, 0)
        Me.Controls.SetChildIndex(Me.lblReportsTo, 0)
        Me.Controls.SetChildIndex(Me.lblTitle, 0)
        Me.Controls.SetChildIndex(Me.lblHireDate, 0)
        Me.Controls.SetChildIndex(Me.lblBirthDate, 0)
        Me.Controls.SetChildIndex(Me.lblExtension, 0)
        Me.Controls.SetChildIndex(Me.lblHomePhone, 0)
        Me.Controls.SetChildIndex(Me.lblCountry, 0)
        Me.Controls.SetChildIndex(Me.lblPostalCode, 0)
        Me.Controls.SetChildIndex(Me.lblRegion, 0)
        Me.Controls.SetChildIndex(Me.lblCity, 0)
        Me.Controls.SetChildIndex(Me.lblAddress, 0)
        Me.Controls.SetChildIndex(Me.lblAvailable, 0)
        Me.Controls.SetChildIndex(Me.lblAssigned, 0)
        Me.Controls.SetChildIndex(Me.lblNotes, 0)
        Me.Controls.SetChildIndex(Me.lblPhoto, 0)
        Me.Controls.SetChildIndex(Me.btnPhoto, 0)
        Me.Controls.SetChildIndex(Me.btnRemove, 0)
        Me.Controls.SetChildIndex(Me.btnAdd, 0)
        Me.Controls.SetChildIndex(Me.rtbNotes, 0)
        Me.Controls.SetChildIndex(Me.btnOK, 0)
        Me.Controls.SetChildIndex(Me.btnCancel, 0)
        Me.Controls.SetChildIndex(Me.btnRules, 0)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmEmployeeEdit_Load(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        Dim strValue As String
        Dim DictEnt As DictionaryEntry
        Dim objEmployeeMgr As EmployeeMgr

        Try
            objEmployeeMgr = objEmployeeMgr.GetInstance
            For Each DictEnt In objEmployeeMgr
                Dim objEmployee As Employee = CType(DictEnt.Value, Employee)
                cboReportsTo.Items.Add(objEmployee)
            Next

            cboReportsTo.Items.Remove(mobjEmployee)

            For Each DictEnt In mobjTerritoryMgr
                Dim objTerritory As Territory = CType(DictEnt.Value, Territory)
                lstAvailable.Items.Add(objTerritory)
            Next

            With mobjEmployee
                For i = 0 To cboCourtesy.Items.Count - 1
                    strValue = CType(cboCourtesy.Items(i), String)
                    If strValue = .TitleOfCourtesy Then
                        cboCourtesy.SelectedIndex = i
                        Exit For
                    End If
                Next

                txtFirstName.Text = .FirstName
                txtLastName.Text = .LastName
                txtTitle.Text = .Title

                If Not .ReportsTo Is Nothing Then
                    For i = 0 To cboReportsTo.Items.Count - 1
                        Dim objEmployee As Employee = _
                        CType(cboReportsTo.Items(i), Employee)
                        If objEmployee.EmployeeID = .ReportsTo.EmployeeID _
                        Then
                            cboReportsTo.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
                dtpBirthDate.Value = .BirthDate
                dtpHireDate.Value = .HireDate
                txtHomePhone.Text = .HomePhone
                txtExtension.Text = .Extension
                txtAddress.Text = .Address
                txtCity.Text = .City
                txtRegion.Text = .Region
                txtPostalCode.Text = .PostalCode
                txtCountry.Text = .Country
                rtbNotes.Text = .Notes

                If Not .Photo Is Nothing Then
                    Dim mStream As New IO.MemoryStream(.Photo)
                    mStream.Write(.Photo, 0, .Photo.Length - 1)
                    picPhoto.Image = Image.FromStream(mStream)
                End If

                For Each DictEnt In .Territories
                    Dim objT As Territory = CType(DictEnt.Value, Territory)
                    lstTerritories.Items.Add(objT)
                    lstAvailable.Items.Remove(objT)
                Next

            End With
        Catch exc As Exception
            LogException(exc)
        End Try

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            If Not lstAvailable.SelectedItem Is Nothing Then
                lstTerritories.Items.Add(lstAvailable.SelectedItem)
                mobjEmployee.Territories.Add(CType(lstAvailable.SelectedItem, _
                Territory))
                lstAvailable.Items.Remove(lstAvailable.SelectedItem)
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnRemove.Click
        Try
            If Not lstTerritories.SelectedItem Is Nothing Then
                Dim objTerritory As Territory = _
                CType(lstTerritories.SelectedItem, Territory)
                lstAvailable.Items.Add(objTerritory)
                mobjEmployee.Territories.Remove(objTerritory.TerritoryID)
                lstTerritories.Items.Remove(lstTerritories.SelectedItem)
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub btnPhoto_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnPhoto.Click
        Dim bytArray() As Byte
        Try
            With OpenFileDialog1
                .Title = "Select Employee Photo"
                .ShowDialog()
            End With

            If OpenFileDialog1.FileName <> "" Then
                Dim fs As New IO.FileStream(OpenFileDialog1.FileName, _
                IO.FileMode.Open)
                ReDim bytArray(Convert.ToInt32(fs.Length - 1))
                fs.Read(bytArray, 0, bytArray.Length - 1)
                mobjEmployee.Photo = bytArray
                fs.Close()

                picPhoto.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub

    Private Sub btnRules_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnRules.Click
        Dim frmRules As New frmBusinessRules(mobjEmployee.GetBusinessRules)
        frmRules.ShowDialog()
        frmRules = Nothing
    End Sub

    Private Sub mobjEmployee_BrokenRule(ByVal IsBroken As Boolean) _
    Handles mobjEmployee.BrokenRule
        If IsBroken Then
            btnOK.Enabled = False
        Else
            btnOK.Enabled = True
        End If
    End Sub

#Region " Validate Events"

    Private Sub txtFirstName_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtFirstName.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.FirstName = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtLastName_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtLastName.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.LastName = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub cboCourtesy_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles cboCourtesy.Validated
        Dim cbo As ComboBox = CType(sender, ComboBox)

        Try
            mobjEmployee.TitleOfCourtesy = cbo.Text
            erpmain.SetError(cbo, "")
        Catch exc As Exception
            erpmain.SetError(cbo, "")
            erpmain.SetError(cbo, exc.Message)
        End Try
    End Sub

    Private Sub txtTitle_TextChanged(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles txtTitle.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.Title = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub cboReportsTo_Click(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles cboReportsTo.Validated
        Dim cbo As ComboBox = CType(sender, ComboBox)

        Try
            mobjEmployee.ReportsTo = CType(cbo.SelectedItem, Employee)
            erpmain.SetError(cbo, "")
        Catch exc As Exception
            erpmain.SetError(cbo, "")
            erpmain.SetError(cbo, exc.Message)
        End Try
    End Sub

    Private Sub dtpBirthDate_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles dtpBirthDate.Validated
        Dim dtp As DateTimePicker = CType(sender, DateTimePicker)

        Try
            mobjEmployee.BirthDate = dtp.Value
            erpmain.SetError(dtp, "")
        Catch exc As Exception
            erpmain.SetError(dtp, "")
            erpmain.SetError(dtp, exc.Message)
        End Try
    End Sub

    Private Sub dtpHireDate_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles dtpHireDate.Validated
        Dim dtp As DateTimePicker = CType(sender, DateTimePicker)

        Try
            mobjEmployee.HireDate = dtp.Value
            erpmain.SetError(dtp, "")
        Catch exc As Exception
            erpmain.SetError(dtp, "")
            erpmain.SetError(dtp, exc.Message)
        End Try
    End Sub

    Private Sub txtHomePhone_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtHomePhone.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.HomePhone = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtExtension_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtExtension.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.Extension = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtAddress_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtAddress.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.Address = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtCountry_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtCountry.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.Country = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtPostalCode_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtPostalCode.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.PostalCode = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtRegion_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtRegion.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.Region = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub txtCity_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles txtCity.Validated
        Dim txt As TextBox = CType(sender, TextBox)

        Try
            mobjEmployee.City = txt.Text
            erpmain.SetError(txt, "")
        Catch exc As Exception
            erpmain.SetError(txt, "")
            erpmain.SetError(txt, exc.Message)
        End Try
    End Sub

    Private Sub rtbNotes_Validated(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles rtbNotes.Validated
        Dim rtb As RichTextBox = CType(sender, RichTextBox)

        Try
            mobjEmployee.Notes = rtb.Text
            erpmain.SetError(rtb, "")
        Catch exc As Exception
            erpmain.SetError(rtb, "")
            erpmain.SetError(rtb, exc.Message)
        End Try
    End Sub

#End Region

    Private Sub btnOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If mobjEmployee.IsDirty Then
                Cursor = Cursors.WaitCursor
                mblnErrors = False
                mobjEmployee.Save()
            End If

            If Not mblnErrors Then
                Close()
            End If
        Catch exc As Exception
            LogException(exc)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub mobjEmployee_Errs(ByVal obj As _
    NorthwindTraders.NorthwindShared.Errors.BusinessErrors) Handles mobjEmployee.Errs
        Try
            Dim i As Integer
            Dim ctl As Control

            For Each ctl In Me.Controls
                If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox _
                Or TypeOf ctl Is RichTextBox Then
                    erpMain.SetError(ctl, "")
                End If
            Next

            For i = 0 To obj.Count - 1
                Select Case obj.Item(i).errProperty
                    Case "Last Name"
                        erpMain.SetError(Me.txtLastName, obj.Item(i).errMessage)
                    Case "First Name"
                        erpMain.SetError(Me.txtFirstName, obj.Item(i).errMessage)
                    Case "Title"
                        erpMain.SetError(Me.txtTitle, obj.Item(i).errMessage)
                    Case "Title Of Courtesy"
                        erpMain.SetError(Me.cboCourtesy, obj.Item(i).errMessage)
                    Case "Birth Date"
                        erpMain.SetError(Me.dtpBirthDate, obj.Item(i).errMessage)
                    Case "Hire Date"
                        erpMain.SetError(Me.dtpHireDate, obj.Item(i).errMessage)
                    Case "Address"
                        erpMain.SetError(Me.txtAddress, obj.Item(i).errMessage)
                    Case "City"
                        erpMain.SetError(Me.txtCity, obj.Item(i).errMessage)
                    Case "Region"
                        erpMain.SetError(Me.txtRegion, obj.Item(i).errMessage)
                    Case "Postal Code"
                        erpMain.SetError(Me.txtPostalCode, obj.Item(i).errMessage)
                    Case "Country"
                        erpMain.SetError(Me.txtCountry, obj.Item(i).errMessage)
                    Case "Home Phone"
                        erpMain.SetError(Me.txtHomePhone, obj.Item(i).errMessage)
                    Case "Extension"
                        erpMain.SetError(Me.txtExtension, obj.Item(i).errMessage)
                    Case "Notes"
                        erpMain.SetError(Me.rtbNotes, obj.Item(i).errMessage)
                    Case "Territories"
                        erpMain.SetError(Me.lstTerritories, obj.Item(i).errMessage)
                End Select
            Next

            mblnErrors = True

        Catch exc As Exception
            logexception(exc)
        End Try
    End Sub

    Private Sub frmEmployeeEdit_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim dlgResult As DialogResult

        Try
            If mobjEmployee.IsDirty Then
                dlgResult = MessageBox.Show("The Employee information has " _
                & "changed, do you want to exit without saving your changes?", _
                "Confirm Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If dlgResult = DialogResult.No Then
                    e.Cancel = True
                Else
                    mobjEmployee.Rollback(mobjTerritoryMgr)
                End If
            End If
        Catch exc As Exception
            LogException(exc)
        End Try
    End Sub
End Class