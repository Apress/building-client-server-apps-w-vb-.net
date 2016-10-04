Option Strict On
Option Explicit On 

Public Class frmEmployeeEdit
    Inherits frmEditBase
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEmployeeEdit))
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.txtCountry = New System.Windows.Forms.TextBox()
        Me.txtPostalCode = New System.Windows.Forms.TextBox()
        Me.cboCourtesy = New System.Windows.Forms.ComboBox()
        Me.cboReportsTo = New System.Windows.Forms.ComboBox()
        Me.dtpBirthDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpHireDate = New System.Windows.Forms.DateTimePicker()
        Me.lstTerritories = New System.Windows.Forms.ListBox()
        Me.lstAvailable = New System.Windows.Forms.ListBox()
        Me.txtRegion = New System.Windows.Forms.TextBox()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.picPhoto = New System.Windows.Forms.PictureBox()
        Me.txtExtension = New System.Windows.Forms.TextBox()
        Me.txtHomePhone = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.btnPhoto = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.rtbNotes = New System.Windows.Forms.RichTextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SuspendLayout()
        '
        'btnRules
        '
        Me.btnRules.Location = New System.Drawing.Point(8, 442)
        Me.btnRules.TabIndex = 22
        Me.btnRules.Visible = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(698, 442)
        Me.btnCancel.TabIndex = 21
        Me.btnCancel.Visible = True
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(610, 442)
        Me.btnOK.TabIndex = 20
        Me.btnOK.Visible = True
        '
        'txtTitle
        '
        Me.erpMain.SetIconAlignment(Me.txtTitle, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtTitle.Location = New System.Drawing.Point(24, 80)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(128, 20)
        Me.txtTitle.TabIndex = 3
        Me.txtTitle.Text = ""
        '
        'txtLastName
        '
        Me.erpMain.SetIconAlignment(Me.txtLastName, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtLastName.Location = New System.Drawing.Point(216, 24)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(136, 20)
        Me.txtLastName.TabIndex = 2
        Me.txtLastName.Text = ""
        '
        'txtFirstName
        '
        Me.erpMain.SetIconAlignment(Me.txtFirstName, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtFirstName.Location = New System.Drawing.Point(104, 24)
        Me.txtFirstName.Name = "txtFirstName"
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
        Me.cboCourtesy.Location = New System.Drawing.Point(24, 24)
        Me.cboCourtesy.Name = "cboCourtesy"
        Me.cboCourtesy.Size = New System.Drawing.Size(72, 21)
        Me.cboCourtesy.TabIndex = 0
        '
        'cboReportsTo
        '
        Me.cboReportsTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.erpMain.SetIconAlignment(Me.cboReportsTo, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.cboReportsTo.Location = New System.Drawing.Point(160, 80)
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
        Me.picPhoto.Location = New System.Drawing.Point(368, 40)
        Me.picPhoto.Name = "picPhoto"
        Me.picPhoto.Size = New System.Drawing.Size(184, 200)
        Me.picPhoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picPhoto.TabIndex = 17
        Me.picPhoto.TabStop = False
        '
        'txtExtension
        '
        Me.erpMain.SetIconAlignment(Me.txtExtension, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.txtExtension.Location = New System.Drawing.Point(160, 192)
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
        'Label1
        '
        Me.erpMain.SetIconAlignment(Me.Label1, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label1.Location = New System.Drawing.Point(24, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Courtesy"
        '
        'Label2
        '
        Me.erpMain.SetIconAlignment(Me.Label2, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label2.Location = New System.Drawing.Point(216, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 16)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Last Name"
        '
        'Label3
        '
        Me.erpMain.SetIconAlignment(Me.Label3, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label3.Location = New System.Drawing.Point(104, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 16)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "First Name"
        '
        'Label4
        '
        Me.erpMain.SetIconAlignment(Me.Label4, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label4.Location = New System.Drawing.Point(160, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 16)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Reports To"
        '
        'Label5
        '
        Me.erpMain.SetIconAlignment(Me.Label5, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label5.Location = New System.Drawing.Point(24, 104)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Title"
        '
        'Label6
        '
        Me.erpMain.SetIconAlignment(Me.Label6, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label6.Location = New System.Drawing.Point(136, 160)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 16)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "Hire Date"
        '
        'Label7
        '
        Me.erpMain.SetIconAlignment(Me.Label7, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label7.Location = New System.Drawing.Point(24, 160)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 16)
        Me.Label7.TabIndex = 27
        Me.Label7.Text = "Birth Date"
        '
        'Label8
        '
        Me.erpMain.SetIconAlignment(Me.Label8, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label8.Location = New System.Drawing.Point(160, 216)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 16)
        Me.Label8.TabIndex = 28
        Me.Label8.Text = "Extension"
        '
        'Label9
        '
        Me.erpMain.SetIconAlignment(Me.Label9, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label9.Location = New System.Drawing.Point(24, 216)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 16)
        Me.Label9.TabIndex = 29
        Me.Label9.Text = "Home Phone"
        '
        'Label10
        '
        Me.erpMain.SetIconAlignment(Me.Label10, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label10.Location = New System.Drawing.Point(16, 408)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 16)
        Me.Label10.TabIndex = 30
        Me.Label10.Text = "Country:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label11
        '
        Me.erpMain.SetIconAlignment(Me.Label11, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label11.Location = New System.Drawing.Point(16, 376)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 16)
        Me.Label11.TabIndex = 31
        Me.Label11.Text = "Postal Code:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label12
        '
        Me.erpMain.SetIconAlignment(Me.Label12, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label12.Location = New System.Drawing.Point(16, 344)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(72, 16)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "Region:"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label13
        '
        Me.erpMain.SetIconAlignment(Me.Label13, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label13.Location = New System.Drawing.Point(16, 312)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 16)
        Me.Label13.TabIndex = 33
        Me.Label13.Text = "City:"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label14
        '
        Me.erpMain.SetIconAlignment(Me.Label14, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label14.Location = New System.Drawing.Point(16, 256)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 16)
        Me.Label14.TabIndex = 34
        Me.Label14.Text = "Address:"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label15
        '
        Me.erpMain.SetIconAlignment(Me.Label15, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label15.Location = New System.Drawing.Point(576, 256)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(112, 16)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "Available Territories"
        '
        'Label16
        '
        Me.erpMain.SetIconAlignment(Me.Label16, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label16.Location = New System.Drawing.Point(416, 256)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(112, 16)
        Me.Label16.TabIndex = 36
        Me.Label16.Text = "Assigned Territories"
        '
        'Label17
        '
        Me.erpMain.SetIconAlignment(Me.Label17, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label17.Location = New System.Drawing.Point(576, 24)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(72, 16)
        Me.Label17.TabIndex = 37
        Me.Label17.Text = "Notes"
        '
        'Label18
        '
        Me.erpMain.SetIconAlignment(Me.Label18, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.Label18.Location = New System.Drawing.Point(368, 24)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 16)
        Me.Label18.TabIndex = 38
        Me.Label18.Text = "Photo"
        '
        'btnPhoto
        '
        Me.erpMain.SetIconAlignment(Me.btnPhoto, System.Windows.Forms.ErrorIconAlignment.MiddleLeft)
        Me.btnPhoto.Image = CType(resources.GetObject("btnPhoto.Image"), System.Drawing.Bitmap)
        Me.btnPhoto.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPhoto.Location = New System.Drawing.Point(528, 16)
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
        Me.rtbNotes.Location = New System.Drawing.Point(576, 40)
        Me.rtbNotes.Name = "rtbNotes"
        Me.rtbNotes.Size = New System.Drawing.Size(192, 200)
        Me.rtbNotes.TabIndex = 15
        Me.rtbNotes.Text = ""
        '
        'frmEmployeeEdit
        '
        Me.AcceptButton = Nothing
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(778, 472)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnRules, Me.btnCancel, Me.btnOK, Me.rtbNotes, Me.btnAdd, Me.btnRemove, Me.btnPhoto, Me.Label18, Me.Label17, Me.Label16, Me.Label15, Me.Label14, Me.Label13, Me.Label12, Me.Label11, Me.Label10, Me.Label9, Me.Label8, Me.Label7, Me.Label6, Me.Label5, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.txtHomePhone, Me.txtExtension, Me.picPhoto, Me.txtCity, Me.txtAddress, Me.txtRegion, Me.lstAvailable, Me.lstTerritories, Me.dtpHireDate, Me.dtpBirthDate, Me.cboReportsTo, Me.cboCourtesy, Me.txtPostalCode, Me.txtCountry, Me.txtFirstName, Me.txtLastName, Me.txtTitle})
        Me.Name = "frmEmployeeEdit"
        Me.Text = "Employee Edit [Details]"
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class