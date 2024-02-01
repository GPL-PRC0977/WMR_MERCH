<%@ Page Title="Log In" Language="VB" MasterPageFile="~/Login.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="TO._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="LoginContent" runat="server">
    <div class="login_div">


     
    <div class="row" style="margin: auto;">

                        <div class="div_logo_holder">
        <h1 style="font-weight: bolder">
            <asp:Image ID="PrimerLogo" runat="server" ImageUrl="~/One Primer Logo.jpg" Height="108px" Width="130px" />
            TRANSFER ORDERS
        </h1>
<%--<center><h4 style="color:red;">FOR TESTING PURPOSES ONLY</h4></center>--%>
                            <asp:Image ID="Image2" CssClass="login_image" runat="server" ImageUrl="~/images/loginimage2.png" />
    </div>

        <div class="col-md-4 login_div2">
            
            
            <p>
                <asp:TextBox ID="textboxUserName" runat="server" CssClass="form-control txt_aign login_user_text" Width="260px" Height="30px" placeholder="Ex: PRIMERGRP\juan.delacruz"></asp:TextBox>
                </p>
            <p>
                <asp:TextBox ID="textboxPassword" runat="server" CssClass="form-control txt_aign login_user_pass" TextMode="Password" Width="260px" Height="30px"></asp:TextBox>
                </p>
            <p>
                <asp:Button ID="buttonLogIn" CssClass="btn btn-primary btn-sm login_btn_" runat="server" Text="Log In" Width="80px" Height="30px" />
            </p>
            <p>
                &nbsp;</p>
                <asp:Label ID="lblStatus" CssClass="invaliduser" runat="server" ForeColor="Red"></asp:Label>
            <p>
                &nbsp;</p>
        </div>
    </div>


    </div>
    
    <p></p>

</asp:Content>
