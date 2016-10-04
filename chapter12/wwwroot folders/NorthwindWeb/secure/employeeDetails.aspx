<%@ Page Language="vb" AutoEventWireup="false" Codebehind="employeeDetails.aspx.vb" Inherits="NorthwindWeb.employeeDetails"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>employeeDetails</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:TextBox id="txtTitle" style="Z-INDEX: 101; LEFT: 32px; POSITION: absolute; TOP: 120px" runat="server"></asp:TextBox>
			<asp:Button id="btnSelect" style="Z-INDEX: 153; LEFT: 600px; POSITION: absolute; TOP: 376px"
				runat="server" Width="24px" Text="<"></asp:Button>
			<asp:Button id="btnRemove" style="Z-INDEX: 152; LEFT: 600px; POSITION: absolute; TOP: 344px"
				runat="server" Width="24px" Text=">"></asp:Button>
			<asp:ListBox id="lstSelected" style="Z-INDEX: 151; LEFT: 456px; POSITION: absolute; TOP: 328px"
				runat="server" Height="106px" Width="128px"></asp:ListBox>
			<asp:ListBox id="lstAvailable" style="Z-INDEX: 150; LEFT: 640px; POSITION: absolute; TOP: 328px"
				runat="server" Height="106px" Width="128px"></asp:ListBox>
			<asp:Label id="lblAvailableTerritories" style="Z-INDEX: 149; LEFT: 640px; POSITION: absolute; TOP: 304px"
				runat="server" Width="128px">Available Territories</asp:Label>
			<asp:Label id="lblSelectedTerritories" style="Z-INDEX: 148; LEFT: 456px; POSITION: absolute; TOP: 304px"
				runat="server">Selected Territories</asp:Label>
			<asp:RegularExpressionValidator id="revCountry" style="Z-INDEX: 147; LEFT: 296px; POSITION: absolute; TOP: 440px"
				runat="server" ErrorMessage="Country cannot be more than 15 characters in length" ControlToValidate="txtCountry"
				ValidationExpression=".{1,15}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revPostal" style="Z-INDEX: 146; LEFT: 224px; POSITION: absolute; TOP: 408px"
				runat="server" ErrorMessage="Postal Code cannot be more than 10 characters in length" ControlToValidate="txtPostal"
				ValidationExpression=".{1,10}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revRegion" style="Z-INDEX: 145; LEFT: 296px; POSITION: absolute; TOP: 376px"
				runat="server" ErrorMessage="Region cannot be more than 15 characters in length" ControlToValidate="txtRegion"
				ValidationExpression=".{1,15}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revCity" style="Z-INDEX: 144; LEFT: 296px; POSITION: absolute; TOP: 344px" runat="server"
				ErrorMessage="City cannot be more than 15 characters in length" ControlToValidate="txtCity" ValidationExpression=".{1,15}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revAddress" style="Z-INDEX: 143; LEFT: 360px; POSITION: absolute; TOP: 296px"
				runat="server" ErrorMessage="Address cannot be more than 60 characters in length" ControlToValidate="txtAddress"
				ValidationExpression=".{1,60}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revExtension" style="Z-INDEX: 142; LEFT: 416px; POSITION: absolute; TOP: 232px"
				runat="server" ErrorMessage="Extension cannot be more than 4 characters in length" ControlToValidate="txtExtension"
				ValidationExpression=".{1,4}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revHomePhone" style="Z-INDEX: 141; LEFT: 192px; POSITION: absolute; TOP: 232px"
				runat="server" ErrorMessage="Home Phone cannot be more than 24 characters in length" ControlToValidate="txtHomePhone"
				ValidationExpression=".{1,24}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revHireDate" style="Z-INDEX: 140; LEFT: 416px; POSITION: absolute; TOP: 176px"
				runat="server" ErrorMessage="Hire Date must be in the format xx/xx/xxxx" ControlToValidate="txtHireDate" ValidationExpression="\d{1,2}/\d{1,2}/\d{4}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revTitle" style="Z-INDEX: 139; LEFT: 192px; POSITION: absolute; TOP: 120px"
				runat="server" ErrorMessage="Title cannot be more than 30 characters in length" ControlToValidate="txtTitle"
				ValidationExpression=".{1,30}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revLastName" style="Z-INDEX: 138; LEFT: 416px; POSITION: absolute; TOP: 72px"
				runat="server" ErrorMessage="Last Name cannot be more than 20 characters in length" ControlToValidate="txtLastName"
				ValidationExpression=".{1,20}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="revFirstNameLength" style="Z-INDEX: 137; LEFT: 224px; POSITION: absolute; TOP: 72px"
				runat="server" ErrorMessage="First Name cannot be more than 10 characters in length" ControlToValidate="txtFirstName"
				ValidationExpression=".{1,10}">*</asp:RegularExpressionValidator>
			<asp:RegularExpressionValidator id="refBirthDate" style="Z-INDEX: 136; LEFT: 192px; POSITION: absolute; TOP: 176px"
				runat="server" ErrorMessage="Birth Date must be in the format xx/xx/xxxx" ControlToValidate="txtBirthDate" ValidationExpression="\d{1,2}/\d{1,2}/\d{4}">*</asp:RegularExpressionValidator>
			<asp:RequiredFieldValidator id="rfvLastName" style="Z-INDEX: 135; LEFT: 416px; POSITION: absolute; TOP: 64px"
				runat="server" ErrorMessage="Last name cannot be null" ControlToValidate="txtLastName">*</asp:RequiredFieldValidator>
			<asp:RequiredFieldValidator id="rfvFirstName" style="Z-INDEX: 134; LEFT: 224px; POSITION: absolute; TOP: 64px"
				runat="server" ErrorMessage="First name cannot be null" ControlToValidate="txtFirstName">*</asp:RequiredFieldValidator>
			<asp:ValidationSummary id="ValidationSummary1" style="Z-INDEX: 133; LEFT: 24px; POSITION: absolute; TOP: 480px"
				runat="server" Width="606px"></asp:ValidationSummary>
			<asp:Button id="btnSave" style="Z-INDEX: 132; LEFT: 720px; POSITION: absolute; TOP: 448px" runat="server"
				Text="Save"></asp:Button>
			<asp:Label id="lblTitle" style="Z-INDEX: 131; LEFT: 32px; POSITION: absolute; TOP: 144px" runat="server">Title</asp:Label>
			<asp:Label id="lblReportsTo" style="Z-INDEX: 130; LEFT: 256px; POSITION: absolute; TOP: 144px"
				runat="server">Reports To</asp:Label>
			<asp:Label id="lblCourtesy" style="Z-INDEX: 129; LEFT: 32px; POSITION: absolute; TOP: 88px"
				runat="server">Courtesy</asp:Label>
			<asp:Label id="lblFirstName" style="Z-INDEX: 128; LEFT: 120px; POSITION: absolute; TOP: 88px"
				runat="server">First Name</asp:Label>
			<asp:Label id="lblLastName" style="Z-INDEX: 127; LEFT: 256px; POSITION: absolute; TOP: 88px"
				runat="server">Last Name</asp:Label>
			<asp:Label id="lblAddress" style="Z-INDEX: 126; LEFT: 32px; POSITION: absolute; TOP: 296px"
				runat="server" Height="11px">Address:</asp:Label>
			<asp:Label id="lblCity" style="Z-INDEX: 125; LEFT: 32px; POSITION: absolute; TOP: 344px" runat="server"
				Height="11px">City:</asp:Label>
			<asp:Label id="lblRegion" style="Z-INDEX: 124; LEFT: 32px; POSITION: absolute; TOP: 376px"
				runat="server" Height="11px">Region:</asp:Label>
			<asp:Label id="lblPostal" style="Z-INDEX: 123; LEFT: 32px; POSITION: absolute; TOP: 408px"
				runat="server" Height="11px">Postal Code:</asp:Label>
			<asp:Label id="lblCountry" style="Z-INDEX: 122; LEFT: 32px; POSITION: absolute; TOP: 440px"
				runat="server" Height="11px">Country:</asp:Label>
			<asp:Label id="lblNotes" style="Z-INDEX: 121; LEFT: 456px; POSITION: absolute; TOP: 64px" runat="server"
				Height="11px">Notes</asp:Label>
			<asp:Label id="Label5" style="Z-INDEX: 120; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="240px" Font-Size="X-Large" Font-Bold="True" Font-Italic="True">Employee Details</asp:Label>
			<asp:Label id="lblBirthDate" style="Z-INDEX: 119; LEFT: 32px; POSITION: absolute; TOP: 200px"
				runat="server">Birth Date</asp:Label>
			<asp:Label id="lblHireDate" style="Z-INDEX: 118; LEFT: 256px; POSITION: absolute; TOP: 200px"
				runat="server">Hire Date</asp:Label>
			<asp:Label id="lblHomePhone" style="Z-INDEX: 117; LEFT: 32px; POSITION: absolute; TOP: 256px"
				runat="server">Home Phone</asp:Label>
			<asp:Label id="lblExtension" style="Z-INDEX: 116; LEFT: 256px; POSITION: absolute; TOP: 256px"
				runat="server">Extension</asp:Label>
			<asp:TextBox id="txtNotes" style="Z-INDEX: 115; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
				Height="200px" Width="192px" TextMode="MultiLine"></asp:TextBox>
			<asp:DropDownList id="ddlCourtesy" style="Z-INDEX: 114; LEFT: 32px; POSITION: absolute; TOP: 64px"
				runat="server">
				<asp:ListItem Value="Mr.">Mr.</asp:ListItem>
				<asp:ListItem Value="Mrs.">Mrs.</asp:ListItem>
				<asp:ListItem Value="Dr.">Dr.</asp:ListItem>
				<asp:ListItem Value="Ms.">Ms.</asp:ListItem>
			</asp:DropDownList>
			<asp:TextBox id="txtLastName" style="Z-INDEX: 113; LEFT: 256px; POSITION: absolute; TOP: 64px"
				runat="server"></asp:TextBox>
			<asp:TextBox id="txtFirstName" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 64px"
				runat="server" Width="104px"></asp:TextBox>
			<asp:TextBox id="txtReportsTo" style="Z-INDEX: 111; LEFT: 256px; POSITION: absolute; TOP: 120px"
				runat="server"></asp:TextBox>
			<asp:TextBox id="txtBirthDate" style="Z-INDEX: 110; LEFT: 32px; POSITION: absolute; TOP: 176px"
				runat="server"></asp:TextBox>
			<asp:TextBox id="txtHireDate" style="Z-INDEX: 109; LEFT: 256px; POSITION: absolute; TOP: 176px"
				runat="server"></asp:TextBox>
			<asp:TextBox id="txtHomePhone" style="Z-INDEX: 108; LEFT: 32px; POSITION: absolute; TOP: 232px"
				runat="server"></asp:TextBox>
			<asp:TextBox id="txtExtension" style="Z-INDEX: 107; LEFT: 256px; POSITION: absolute; TOP: 232px"
				runat="server"></asp:TextBox>
			<asp:TextBox id="txtAddress" style="Z-INDEX: 106; LEFT: 136px; POSITION: absolute; TOP: 296px"
				runat="server" Height="40px" Width="224px" TextMode="MultiLine"></asp:TextBox>
			<asp:TextBox id="txtCity" style="Z-INDEX: 105; LEFT: 136px; POSITION: absolute; TOP: 344px" runat="server"></asp:TextBox>
			<asp:TextBox id="txtRegion" style="Z-INDEX: 104; LEFT: 136px; POSITION: absolute; TOP: 376px"
				runat="server"></asp:TextBox>
			<asp:TextBox id="txtPostal" style="Z-INDEX: 103; LEFT: 136px; POSITION: absolute; TOP: 408px"
				runat="server" Width="80px"></asp:TextBox>
			<asp:TextBox id="txtCountry" style="Z-INDEX: 102; LEFT: 136px; POSITION: absolute; TOP: 440px"
				runat="server"></asp:TextBox>
		</form>
	</body>
</HTML>
