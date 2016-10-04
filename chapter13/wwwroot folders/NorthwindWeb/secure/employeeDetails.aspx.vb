Option Explicit On 
Option Strict On

Imports NorthwindTraders.NorthwindShared.Interfaces
Imports NorthwindTraders.NorthwindShared.Structures
Imports NorthwindTraders.NorthwindShared.Errors
Imports BusinessRules.Errors

Public Class employeeDetails
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents txtTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCountry As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtPostal As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtRegion As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCity As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtAddress As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtExtension As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtHomePhone As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtHireDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtBirthDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtReportsTo As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtFirstName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtLastName As System.Web.UI.WebControls.TextBox
    Protected WithEvents ddlCourtesy As System.Web.UI.WebControls.DropDownList
    Protected WithEvents txtNotes As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblExtension As System.Web.UI.WebControls.Label
    Protected WithEvents lblHomePhone As System.Web.UI.WebControls.Label
    Protected WithEvents lblHireDate As System.Web.UI.WebControls.Label
    Protected WithEvents lblBirthDate As System.Web.UI.WebControls.Label
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents lblNotes As System.Web.UI.WebControls.Label
    Protected WithEvents lblCountry As System.Web.UI.WebControls.Label
    Protected WithEvents lblPostal As System.Web.UI.WebControls.Label
    Protected WithEvents lblRegion As System.Web.UI.WebControls.Label
    Protected WithEvents lblCity As System.Web.UI.WebControls.Label
    Protected WithEvents lblAddress As System.Web.UI.WebControls.Label
    Protected WithEvents lblLastName As System.Web.UI.WebControls.Label
    Protected WithEvents lblFirstName As System.Web.UI.WebControls.Label
    Protected WithEvents lblCourtesy As System.Web.UI.WebControls.Label
    Protected WithEvents lblReportsTo As System.Web.UI.WebControls.Label
    Protected WithEvents lblTitle As System.Web.UI.WebControls.Label
    Protected WithEvents rfvFirstName As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents ValidationSummary1 As System.Web.UI.WebControls.ValidationSummary
    Protected WithEvents btnSave As System.Web.UI.WebControls.Button
    Protected WithEvents rfvLastName As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents revFirstNameLength As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents refBirthDate As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revLastName As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revTitle As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revHireDate As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revHomePhone As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revExtension As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revAddress As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revCity As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revRegion As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revPostal As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents revCountry As System.Web.UI.WebControls.RegularExpressionValidator
    Protected WithEvents lblSelectedTerritories As System.Web.UI.WebControls.Label
    Protected WithEvents lblAvailableTerritories As System.Web.UI.WebControls.Label
    Protected WithEvents lstAvailable As System.Web.UI.WebControls.ListBox
    Protected WithEvents lstSelected As System.Web.UI.WebControls.ListBox
    Protected WithEvents btnRemove As System.Web.UI.WebControls.Button
    Protected WithEvents btnSelect As System.Web.UI.WebControls.Button

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objService As IEmployee
        Dim sEmp As structEmployee
        Dim i As Integer
        Dim intID As Integer
        Dim dsTerritories As DataSet
        Dim dRow As DataRow

        If Not Page.IsPostBack Then
            Dim locRes As New Resources.ResourceManager("NorthwindWeb", _
            GetType(NorthwindWeb.employeeDetails).Module.Assembly)

            Me.lblCourtesy.Text = locRes.GetString("COURTESY")
            Me.lblFirstName.Text = locRes.GetString("FIRST_NAME")
            Me.lblLastName.Text = locRes.GetString("LAST_NAME")
            Me.lblTitle.Text = locRes.GetString("TITLE")
            Me.lblReportsTo.Text = locRes.GetString("REPORTS_TO")
            Me.lblBirthDate.Text = locRes.GetString("BIRTH_DATE")
            Me.lblHireDate.Text = locRes.GetString("HIRE_DATE")
            Me.lblHomePhone.Text = locRes.GetString("HOME_PHONE")
            Me.lblExtension.Text = locRes.GetString("EXTENSION")
            Me.lblAddress.Text = locRes.GetString("ADDRESS")
            Me.lblCity.Text = locRes.GetString("CITY")
            Me.lblRegion.Text = locRes.GetString("REGION")
            Me.lblPostal.Text = locRes.GetString("POSTAL_CODE")
            Me.lblCountry.Text = locRes.GetString("COUNTRY")
            Me.lblNotes.Text = locRes.GetString("NOTES")
            Me.lblSelectedTerritories.Text = locRes.GetString("SELECTED_TER")
            Me.lblAvailableTerritories.Text = locRes.GetString("AVAILABLE_TER")
            Me.btnSave.Text = locRes.GetString("OK")

            Dim objTerritories As ITerritory

            objTerritories = CType(Activator.GetObject(GetType(ITerritory), _
            "http://localhost:80/northwind/TerritoryDC.rem"), ITerritory)

            dsTerritories = objTerritories.LoadProxy
            objTerritories = Nothing

            For Each dRow In dsTerritories.Tables(0).Rows
                Dim strDesc As String = _
                Convert.ToString(dRow.Item("TerritoryDescription"))
                Dim strID As String = Convert.ToString(dRow.Item("TerritoryID"))
                Dim lst As New ListItem(strDesc)
                lst.Value = strID
                lstAvailable.Items.Add(lst)
            Next

            If Page.Request.QueryString.Item("ID") <> "0" Then
                objService = CType(Activator.GetObject(GetType(IEmployee), _
                "http://localhost:80/northwind/EmployeeDC.rem"), IEmployee)
                intID = Convert.ToInt32(Page.Request.QueryString.Item("ID"))
                sEmp = objService.LoadRecord(intID)
                objService = Nothing

                With sEmp
                    For i = 0 To ddlCourtesy.Items.Count - 1
                        If ddlCourtesy.Items(i).Text = .TitleOfCourtesy Then
                            ddlCourtesy.SelectedIndex = i
                            Exit For
                        End If
                    Next

                    txtFirstName.Text = .FirstName
                    txtLastName.Text = .LastName
                    txtTitle.Text = .Title
                    txtReportsTo.Text = .ReportsToFirstName & " " & .ReportsToLastName
                    txtBirthDate.Text = .BirthDate.ToString("D")
                    txtHireDate.Text = .HireDate.ToString("D")
                    txtHomePhone.Text = .HomePhone
                    txtExtension.Text = .Extension
                    txtAddress.Text = .Address
                    txtCity.Text = .City
                    txtRegion.Text = .Region
                    txtPostal.Text = .PostalCode
                    txtCountry.Text = .Country
                    txtNotes.Text = .Notes

                    For i = 0 To .Territories.Length - 1
                        For Each dRow In dsTerritories.Tables(0).Rows
                            Dim strDesc As String = _
                            Convert.ToString(dRow.Item("TerritoryDescription"))
                            Dim strID As String = Convert.ToString(dRow.Item("TerritoryID"))

                            If strID = .Territories(i) Then
                                Dim lst As New ListItem(strDesc)
                                lst.Value = strID
                                lstSelected.Items.Add(lst)
                                lstAvailable.Items.Remove(lst)
                                Exit For
                            End If
                        Next
                    Next

                End With
            End If
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim s As structEmployee
        Dim objEmployee As IEmployee
        Dim intID As Integer
        Dim objBusErr As BusinessErrors

        With s
            .EmployeeID = Convert.ToInt32(Page.Request.QueryString.Item("ID"))
            .TitleOfCourtesy = ddlCourtesy.SelectedItem.Text
            .FirstName = txtFirstName.Text
            .LastName = txtLastName.Text
            .Title = txtTitle.Text
            If txtBirthDate.Text.Length > 0 Then
                .BirthDate = Convert.ToDateTime(txtBirthDate.Text)
            End If
            If txtHireDate.Text.Length > 0 Then
                .HireDate = Convert.ToDateTime(txtHireDate.Text)
            End If
            .HomePhone = txtHomePhone.Text
            .Extension = txtExtension.Text
            .Address = txtAddress.Text
            .City = txtCity.Text
            .Region = txtRegion.Text
            .PostalCode = txtPostal.Text
            .Country = txtCountry.Text
            .Notes = txtNotes.Text

            Dim i As Integer
            ReDim .Territories(lstSelected.Items.Count - 1)

            For i = 0 To lstSelected.Items.Count - 1
                .Territories(i) = lstSelected.Items(i).Value
            Next
        End With

        objEmployee = CType(Activator.GetObject(GetType(IEmployee), _
            "http://localhost:80/northwind/EmployeeDC.rem"), IEmployee)
        objBusErr = objEmployee.Save(s, intID)
        objEmployee = Nothing

        If objBusErr Is Nothing Then
            Response.Redirect("Employees.aspx")
        Else
            ShowErrors(objBusErr)
        End If

    End Sub

    Private Sub ShowErrors(ByVal objBE As BusinessErrors)
        Dim i As Integer

        For i = 0 To objBE.Count - 1
            Select Case objBE.Item(i).errProperty
                Case "Last Name"
                    txtLastName.BackColor = Color.LightPink
                    txtLastName.ToolTip = objBE.Item(i).errMessage
                Case "First Name"
                    txtFirstName.BackColor = Color.LightPink
                    txtFirstName.ToolTip = objBE.Item(i).errMessage
                Case "Title"
                    txtTitle.BackColor = Color.LightPink
                    txtTitle.ToolTip = objBE.Item(i).errMessage
                Case "Title Of Courtesy"
                    ddlCourtesy.BackColor = Color.LightPink
                    ddlCourtesy.ToolTip = objBE.Item(i).errMessage
                Case "Birth Date"
                    txtBirthDate.BackColor = Color.LightPink
                    txtBirthDate.ToolTip = objBE.Item(i).errMessage
                Case "Hire Date"
                    txtHireDate.BackColor = Color.LightPink
                    txtHireDate.ToolTip = objBE.Item(i).errMessage
                Case "Address"
                    txtAddress.BackColor = Color.LightPink
                    txtAddress.ToolTip = objBE.Item(i).errMessage
                Case "City"
                    txtCity.BackColor = Color.LightPink
                    txtCity.ToolTip = objBE.Item(i).errMessage
                Case "Region"
                    txtRegion.BackColor = Color.LightPink
                    txtRegion.ToolTip = objBE.Item(i).errMessage
                Case "Postal Code"
                    txtPostal.BackColor = Color.LightPink
                    txtPostal.ToolTip = objBE.Item(i).errMessage
                Case "Country"
                    txtCountry.BackColor = Color.LightPink
                    txtCountry.ToolTip = objBE.Item(i).errMessage
                Case "Home Phone"
                    txtHomePhone.BackColor = Color.LightPink
                    txtHomePhone.ToolTip = objBE.Item(i).errMessage
                Case "Extension"
                    txtExtension.BackColor = Color.LightPink
                    txtExtension.ToolTip = objBE.Item(i).errMessage
                Case "Notes"
                    txtNotes.BackColor = Color.LightPink
                    txtNotes.ToolTip = objBE.Item(i).errMessage
                Case "Territories"
                    'Nothing yet
            End Select
        Next
    End Sub

    Private Sub btnSelect_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnSelect.Click
        Dim lst As ListItem = lstAvailable.SelectedItem
        lstAvailable.Items.Remove(lst)
        lstSelected.Items.Add(lst)
        lstSelected.SelectedIndex = lstSelected.Items.Count - 1
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnRemove.Click
        Dim lst As ListItem = lstSelected.SelectedItem
        lstSelected.Items.Remove(lst)
        lstAvailable.Items.Add(lst)
        lstAvailable.SelectedIndex = 0
    End Sub

End Class
