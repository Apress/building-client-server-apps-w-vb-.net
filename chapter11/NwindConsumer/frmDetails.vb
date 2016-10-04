Option Strict On
Option Explicit On 

Public Class frmDetails
    Inherits System.Windows.Forms.Form

    Private msEmployee As localhost.structEmployee

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    Public Sub New(ByVal sEmployee As localhost.structEmployee)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        msEmployee = sEmployee
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
    Friend WithEvents lblAvailable As System.Windows.Forms.Label
    Friend WithEvents rtbNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnPhoto As System.Windows.Forms.Button
    Friend WithEvents lblPhoto As System.Windows.Forms.Label
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents lblAssigned As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblRegion As System.Windows.Forms.Label
    Friend WithEvents lblCountry As System.Windows.Forms.Label
    Friend WithEvents lblHomePhone As System.Windows.Forms.Label
    Friend WithEvents lblExtension As System.Windows.Forms.Label
    Friend WithEvents lblBirthDate As System.Windows.Forms.Label
    Friend WithEvents lblHireDate As System.Windows.Forms.Label
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblReportsTo As System.Windows.Forms.Label
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents lblCourtesy As System.Windows.Forms.Label
    Friend WithEvents txtHomePhone As System.Windows.Forms.TextBox
    Friend WithEvents txtExtension As System.Windows.Forms.TextBox
    Friend WithEvents picPhoto As System.Windows.Forms.PictureBox
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtRegion As System.Windows.Forms.TextBox
    Friend WithEvents lstAvailable As System.Windows.Forms.ListBox
    Friend WithEvents lstTerritories As System.Windows.Forms.ListBox
    Friend WithEvents dtpHireDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpBirthDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboReportsTo As System.Windows.Forms.ComboBox
    Friend WithEvents cboCourtesy As System.Windows.Forms.ComboBox
    Friend WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Friend WithEvents txtCountry As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDetails))
        Me.lblAvailable = New System.Windows.Forms.Label
        Me.rtbNotes = New System.Windows.Forms.RichTextBox
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnRemove = New System.Windows.Forms.Button
        Me.btnPhoto = New System.Windows.Forms.Button
        Me.lblPhoto = New System.Windows.Forms.Label
        Me.lblNotes = New System.Windows.Forms.Label
        Me.lblAssigned = New System.Windows.Forms.Label
        Me.lblAddress = New System.Windows.Forms.Label
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblRegion = New System.Windows.Forms.Label
        Me.lblCountry = New System.Windows.Forms.Label
        Me.lblHomePhone = New System.Windows.Forms.Label
        Me.lblExtension = New System.Windows.Forms.Label
        Me.lblBirthDate = New System.Windows.Forms.Label
        Me.lblHireDate = New System.Windows.Forms.Label
        Me.lblTitle = New System.Windows.Forms.Label
        Me.lblReportsTo = New System.Windows.Forms.Label
        Me.lblFirstName = New System.Windows.Forms.Label
        Me.lblLastName = New System.Windows.Forms.Label
        Me.lblCourtesy = New System.Windows.Forms.Label
        Me.txtHomePhone = New System.Windows.Forms.TextBox
        Me.txtExtension = New System.Windows.Forms.TextBox
        Me.picPhoto = New System.Windows.Forms.PictureBox
        Me.txtCity = New System.Windows.Forms.TextBox
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.txtRegion = New System.Windows.Forms.TextBox
        Me.lstAvailable = New System.Windows.Forms.ListBox
        Me.lstTerritories = New System.Windows.Forms.ListBox
        Me.dtpHireDate = New System.Windows.Forms.DateTimePicker
        Me.dtpBirthDate = New System.Windows.Forms.DateTimePicker
        Me.cboReportsTo = New System.Windows.Forms.ComboBox
        Me.cboCourtesy = New System.Windows.Forms.ComboBox
        Me.txtPostalCode = New System.Windows.Forms.TextBox
        Me.txtCountry = New System.Windows.Forms.TextBox
        Me.txtFirstName = New System.Windows.Forms.TextBox
        Me.txtLastName = New System.Windows.Forms.TextBox
        Me.txtTitle = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'lblAvailable
        '
        Me.lblAvailable.Location = New System.Drawing.Point(576, 256)
        Me.lblAvailable.Name = "lblAvailable"
        Me.lblAvailable.Size = New System.Drawing.Size(112, 16)
        Me.lblAvailable.TabIndex = 77
        Me.lblAvailable.Text = "Available Territories"
        '
        'rtbNotes
        '
        Me.rtbNotes.Location = New System.Drawing.Point(584, 40)
        Me.rtbNotes.Name = "rtbNotes"
        Me.rtbNotes.Size = New System.Drawing.Size(184, 184)
        Me.rtbNotes.TabIndex = 55
        Me.rtbNotes.Text = ""
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(544, 312)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(24, 23)
        Me.btnAdd.TabIndex = 58
        Me.btnAdd.Text = "<"
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(544, 352)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(24, 23)
        Me.btnRemove.TabIndex = 59
        Me.btnRemove.Text = ">"
        '
        'btnPhoto
        '
        Me.btnPhoto.Image = CType(resources.GetObject("btnPhoto.Image"), System.Drawing.Image)
        Me.btnPhoto.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPhoto.Location = New System.Drawing.Point(536, 16)
        Me.btnPhoto.Name = "btnPhoto"
        Me.btnPhoto.Size = New System.Drawing.Size(24, 23)
        Me.btnPhoto.TabIndex = 54
        '
        'lblPhoto
        '
        Me.lblPhoto.Location = New System.Drawing.Point(384, 24)
        Me.lblPhoto.Name = "lblPhoto"
        Me.lblPhoto.Size = New System.Drawing.Size(72, 16)
        Me.lblPhoto.TabIndex = 76
        Me.lblPhoto.Text = "Photo"
        '
        'lblNotes
        '
        Me.lblNotes.Location = New System.Drawing.Point(584, 24)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(72, 16)
        Me.lblNotes.TabIndex = 75
        Me.lblNotes.Text = "Notes"
        '
        'lblAssigned
        '
        Me.lblAssigned.Location = New System.Drawing.Point(416, 256)
        Me.lblAssigned.Name = "lblAssigned"
        Me.lblAssigned.Size = New System.Drawing.Size(112, 16)
        Me.lblAssigned.TabIndex = 74
        Me.lblAssigned.Text = "Assigned Territories"
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(16, 256)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(72, 16)
        Me.lblAddress.TabIndex = 73
        Me.lblAddress.Text = "Address:"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(16, 312)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(72, 16)
        Me.lblCity.TabIndex = 72
        Me.lblCity.Text = "City:"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblRegion
        '
        Me.lblRegion.Location = New System.Drawing.Point(16, 344)
        Me.lblRegion.Name = "lblRegion"
        Me.lblRegion.Size = New System.Drawing.Size(72, 16)
        Me.lblRegion.TabIndex = 71
        Me.lblRegion.Text = "Region:"
        Me.lblRegion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCountry
        '
        Me.lblCountry.Location = New System.Drawing.Point(16, 408)
        Me.lblCountry.Name = "lblCountry"
        Me.lblCountry.Size = New System.Drawing.Size(72, 16)
        Me.lblCountry.TabIndex = 70
        Me.lblCountry.Text = "Country:"
        Me.lblCountry.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblHomePhone
        '
        Me.lblHomePhone.Location = New System.Drawing.Point(24, 216)
        Me.lblHomePhone.Name = "lblHomePhone"
        Me.lblHomePhone.Size = New System.Drawing.Size(72, 16)
        Me.lblHomePhone.TabIndex = 69
        Me.lblHomePhone.Text = "Home Phone"
        '
        'lblExtension
        '
        Me.lblExtension.Location = New System.Drawing.Point(176, 216)
        Me.lblExtension.Name = "lblExtension"
        Me.lblExtension.Size = New System.Drawing.Size(64, 16)
        Me.lblExtension.TabIndex = 68
        Me.lblExtension.Text = "Extension"
        '
        'lblBirthDate
        '
        Me.lblBirthDate.Location = New System.Drawing.Point(24, 160)
        Me.lblBirthDate.Name = "lblBirthDate"
        Me.lblBirthDate.Size = New System.Drawing.Size(72, 16)
        Me.lblBirthDate.TabIndex = 67
        Me.lblBirthDate.Text = "Birth Date"
        '
        'lblHireDate
        '
        Me.lblHireDate.Location = New System.Drawing.Point(136, 160)
        Me.lblHireDate.Name = "lblHireDate"
        Me.lblHireDate.Size = New System.Drawing.Size(72, 16)
        Me.lblHireDate.TabIndex = 66
        Me.lblHireDate.Text = "Hire Date"
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(16, 104)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(128, 16)
        Me.lblTitle.TabIndex = 65
        Me.lblTitle.Text = "Title"
        '
        'lblReportsTo
        '
        Me.lblReportsTo.Location = New System.Drawing.Point(168, 104)
        Me.lblReportsTo.Name = "lblReportsTo"
        Me.lblReportsTo.Size = New System.Drawing.Size(152, 16)
        Me.lblReportsTo.TabIndex = 64
        Me.lblReportsTo.Text = "Reports To"
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(112, 48)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(88, 16)
        Me.lblFirstName.TabIndex = 63
        Me.lblFirstName.Text = "First Name"
        '
        'lblLastName
        '
        Me.lblLastName.Location = New System.Drawing.Point(232, 48)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(128, 16)
        Me.lblLastName.TabIndex = 62
        Me.lblLastName.Text = "Last Name"
        '
        'lblCourtesy
        '
        Me.lblCourtesy.Location = New System.Drawing.Point(16, 48)
        Me.lblCourtesy.Name = "lblCourtesy"
        Me.lblCourtesy.Size = New System.Drawing.Size(72, 16)
        Me.lblCourtesy.TabIndex = 61
        Me.lblCourtesy.Text = "Courtesy"
        '
        'txtHomePhone
        '
        Me.txtHomePhone.Location = New System.Drawing.Point(24, 192)
        Me.txtHomePhone.Name = "txtHomePhone"
        Me.txtHomePhone.Size = New System.Drawing.Size(120, 20)
        Me.txtHomePhone.TabIndex = 47
        Me.txtHomePhone.Text = ""
        '
        'txtExtension
        '
        Me.txtExtension.Location = New System.Drawing.Point(176, 192)
        Me.txtExtension.Name = "txtExtension"
        Me.txtExtension.Size = New System.Drawing.Size(64, 20)
        Me.txtExtension.TabIndex = 48
        Me.txtExtension.Text = ""
        '
        'picPhoto
        '
        Me.picPhoto.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picPhoto.Location = New System.Drawing.Point(384, 40)
        Me.picPhoto.Name = "picPhoto"
        Me.picPhoto.Size = New System.Drawing.Size(176, 184)
        Me.picPhoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picPhoto.TabIndex = 57
        Me.picPhoto.TabStop = False
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(112, 304)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(168, 20)
        Me.txtCity.TabIndex = 50
        Me.txtCity.Text = ""
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(112, 256)
        Me.txtAddress.Multiline = True
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(248, 40)
        Me.txtAddress.TabIndex = 49
        Me.txtAddress.Text = ""
        '
        'txtRegion
        '
        Me.txtRegion.Location = New System.Drawing.Point(112, 336)
        Me.txtRegion.Name = "txtRegion"
        Me.txtRegion.Size = New System.Drawing.Size(168, 20)
        Me.txtRegion.TabIndex = 51
        Me.txtRegion.Text = ""
        '
        'lstAvailable
        '
        Me.lstAvailable.Location = New System.Drawing.Point(576, 272)
        Me.lstAvailable.Name = "lstAvailable"
        Me.lstAvailable.Size = New System.Drawing.Size(120, 147)
        Me.lstAvailable.Sorted = True
        Me.lstAvailable.TabIndex = 60
        '
        'lstTerritories
        '
        Me.lstTerritories.Location = New System.Drawing.Point(416, 272)
        Me.lstTerritories.Name = "lstTerritories"
        Me.lstTerritories.Size = New System.Drawing.Size(120, 147)
        Me.lstTerritories.Sorted = True
        Me.lstTerritories.TabIndex = 56
        '
        'dtpHireDate
        '
        Me.dtpHireDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpHireDate.Location = New System.Drawing.Point(136, 136)
        Me.dtpHireDate.Name = "dtpHireDate"
        Me.dtpHireDate.Size = New System.Drawing.Size(88, 20)
        Me.dtpHireDate.TabIndex = 46
        '
        'dtpBirthDate
        '
        Me.dtpBirthDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpBirthDate.Location = New System.Drawing.Point(24, 136)
        Me.dtpBirthDate.Name = "dtpBirthDate"
        Me.dtpBirthDate.Size = New System.Drawing.Size(88, 20)
        Me.dtpBirthDate.TabIndex = 45
        '
        'cboReportsTo
        '
        Me.cboReportsTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReportsTo.Location = New System.Drawing.Point(168, 80)
        Me.cboReportsTo.Name = "cboReportsTo"
        Me.cboReportsTo.Size = New System.Drawing.Size(152, 21)
        Me.cboReportsTo.TabIndex = 44
        '
        'cboCourtesy
        '
        Me.cboCourtesy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCourtesy.Items.AddRange(New Object() {"Mr.", "Mrs.", "Dr.", "Ms."})
        Me.cboCourtesy.Location = New System.Drawing.Point(16, 24)
        Me.cboCourtesy.Name = "cboCourtesy"
        Me.cboCourtesy.Size = New System.Drawing.Size(72, 21)
        Me.cboCourtesy.TabIndex = 40
        '
        'txtPostalCode
        '
        Me.txtPostalCode.Location = New System.Drawing.Point(112, 368)
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.Size = New System.Drawing.Size(168, 20)
        Me.txtPostalCode.TabIndex = 52
        Me.txtPostalCode.Text = ""
        '
        'txtCountry
        '
        Me.txtCountry.Location = New System.Drawing.Point(112, 400)
        Me.txtCountry.Name = "txtCountry"
        Me.txtCountry.Size = New System.Drawing.Size(168, 20)
        Me.txtCountry.TabIndex = 53
        Me.txtCountry.Text = ""
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(112, 24)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(88, 20)
        Me.txtFirstName.TabIndex = 41
        Me.txtFirstName.Text = ""
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(232, 24)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(128, 20)
        Me.txtLastName.TabIndex = 42
        Me.txtLastName.Text = ""
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(16, 80)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(128, 20)
        Me.txtTitle.TabIndex = 43
        Me.txtTitle.Text = ""
        '
        'frmDetails
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 462)
        Me.Controls.Add(Me.lblAvailable)
        Me.Controls.Add(Me.rtbNotes)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.btnPhoto)
        Me.Controls.Add(Me.lblPhoto)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.lblAssigned)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblRegion)
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
        Me.Name = "frmDetails"
        Me.Text = "frmDetails"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmDetails_Load(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        Dim strValue As String

        Try
            With msEmployee
                cboCourtesy.Text = .TitleOfCourtesy
                txtFirstName.Text = .FirstName
                txtLastName.Text = .LastName
                txtTitle.Text = .Title
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

            End With
        Catch exc As Exception
            MessageBox.Show(exc.Message)
        End Try
    End Sub

End Class
