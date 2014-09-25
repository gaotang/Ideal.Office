<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Ideal.Office.Web._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
     <asp:Button ID="btnExcelToTable" runat="server" Text="ExcelToTable" OnClick="btnExcelToTable_Click" />
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    <h3>ExcelTable Data:</h3>
        <%=excelContent %>
</asp:Content>
