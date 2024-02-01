<%@ Page Title="Main" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Main.aspx.vb" Inherits="TO.Main" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div>
        <h1 style="font-weight: bolder">
            <asp:Image ID="PrimerLogo" runat="server" ImageUrl="~/One Primer Logo.jpg" Height="108px" Width="130px" />
            &nbsp;TRANSFER ORDERS
        </h1>
    </div>

    <br /><br />

    <div class="row">
        
        <div class="col-md-9">
            <ul class="nav navbar-nav">
                <li><a runat="server" href="~/MR.aspx"><h3>Merchandise Replenishment</h3></a></li>
            </ul>
        </div>
        <div class="col-md-9">
            <ul class="nav navbar-nav">
                <%--<li><asp:HyperLink ID="barcodeslink" runat="server" href="~/Barcode.aspx"><h4>Barcode Types</h4></asp:HyperLink></li>
                <li><asp:HyperLink ID="userslink" runat="server" href="~/Users.aspx"><h4>Users</h4></asp:HyperLink></li>--%>  
                <%--<li><a runat="server" href="~/Default.aspx"><h4>Log Out</h4></a></li>--%>
            </ul>    
        </div>
    </div>
    
    <p></p>
    
</asp:Content>
