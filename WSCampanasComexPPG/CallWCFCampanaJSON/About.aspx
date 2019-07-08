<%@ Page Title="About" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="About.aspx.cs" Inherits="CallWCFCampanaJSON.About" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <img src="Image/comex_layout_2.png" style="float:left; vertical-align:top;" width="120" height="30"> 
    <img src="Image/PPG_Layout.png" style="float:right; vertical-align:top;" width="70" height="30">

    <h2><%: Title %>.</h2>
    <h3>Your application description page.</h3>
    <p>Use this area to provide additional information.</p>
    <img src="Image/PPG_Layout.png" />
    <asp:Button runat="server" Text="click" OnClick="Unnamed1_Click" />
    
</asp:Content>


