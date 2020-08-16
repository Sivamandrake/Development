<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="Headtrax._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <br />
    <asp:label id="lblAlias" title="Alias" runat ="server">Alias</asp:label>
   <asp:TextBox ID="txtAlias" runat ="server"></asp:TextBox>
    <asp:Button ID="btnSubmit" Text="Submit" runat="server" />
    <asp:ListBox ID="list" runat="server" Height="344px" Width="939px"></asp:ListBox>
</asp:Content>
