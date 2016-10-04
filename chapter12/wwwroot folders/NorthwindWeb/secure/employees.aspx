<%@ Page Language="vb" AutoEventWireup="false" Codebehind="employees.aspx.vb" Inherits="NorthwindWeb.employees"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>employees</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:DataGrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 80px; POSITION: absolute; TOP: 136px"
				runat="server" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" BackColor="White"
				CellPadding="3" GridLines="Vertical" AutoGenerateColumns="False">
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#008A8C"></SelectedItemStyle>
				<AlternatingItemStyle BackColor="#DCDCDC"></AlternatingItemStyle>
				<ItemStyle ForeColor="Black" BackColor="#EEEEEE"></ItemStyle>
				<HeaderStyle Font-Bold="True" ForeColor="White" BackColor="#000084"></HeaderStyle>
				<FooterStyle ForeColor="Black" BackColor="#CCCCCC"></FooterStyle>
				<Columns>
					<asp:HyperLinkColumn Text="Details" DataNavigateUrlField="EmployeeID" DataNavigateUrlFormatString="EmployeeDetails.aspx?ID={0}"></asp:HyperLinkColumn>
					<asp:BoundColumn DataField="FirstName" HeaderText="First Name"></asp:BoundColumn>
					<asp:BoundColumn DataField="LastName" HeaderText="Last Name"></asp:BoundColumn>
					<asp:BoundColumn DataField="Title" HeaderText="Title"></asp:BoundColumn>
				</Columns>
				<PagerStyle HorizontalAlign="Center" ForeColor="Black" BackColor="#999999" Mode="NumericPages"></PagerStyle>
			</asp:DataGrid>
			<asp:HyperLink id="hypAdd" style="Z-INDEX: 103; LEFT: 80px; POSITION: absolute; TOP: 96px" runat="server"
				Width="96px" NavigateUrl="EmployeeDetails.aspx?ID=0">Add Employee</asp:HyperLink>
			<asp:Label id="Label1" style="Z-INDEX: 102; LEFT: 16px; POSITION: absolute; TOP: 16px" runat="server"
				Width="272px" Height="48px" Font-Size="X-Large" Font-Bold="True" Font-Italic="True">Employee List</asp:Label>
		</form>
	</body>
</HTML>
