<%@ Page Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="Upload.aspx.vb" Inherits="TO.Upload" %>

<asp:Content ID="MRContent" ContentPlaceHolderID="MainContent" runat="server">
    <style>
        .upload_tbl{
            border: solid 1px lightgray;
            width: 100%;
        }
        .upload_tbl td{
            padding: 10px;
        }
        .upload_holder{
            position: absolute;
            padding: 10px;
            width: 97%;
            margin: auto;
            text-align: center;
        }
        .table_holder{
            border: solid 1px;
            padding: 20px;
            width: 500px;
            margin: auto;
            border: solid 1px lightgray;
        }
        .login_imageMR2{
            position: absolute;
            margin-left: -230px;
            margin-top: 20px;
            margin-right: 20px;
        }
    </style>
        <div class="MR_nav">
        <asp:Image ID="Image2" CssClass="login_imageMR" runat="server" ImageUrl="~/images/Properties.png" />
        <asp:Label ID="sitename" CssClass="site_name" runat="server" Text="Transfer Orders"></asp:Label>

        <asp:Button ID="logout_btn" Text="Logout" runat="server" CssClass="logoutbtn_css" />
<%--        <asp:Button ID="exportbtn" Text="Export to Excel" runat="server" CssClass="exportbtn_css" />--%>
        <asp:Button ID="homebtn" Text="Home" runat="server" CssClass="homebtn_css" />
        
    </div>
    <div class="MR_nav2">
        <asp:Label ID="activeUser" runat="server" ForeColor="gray"></asp:Label>
        <asp:Label ID="lbldate" CssClass="lbldate_css" runat="server"><%= DateTime.Now %></asp:Label>
        
    </div>
    <div class="upload_holder">

            
        <div class="table_holder">
            <asp:Image ID="Image1" CssClass="login_imageMR2" runat="server" ImageUrl="~/images/Open.png" />
            <h3>Document No: <asp:Label ID="selectedDocNo" runat="server" ForeColor="Blue"></asp:Label></h3><br />
                <table class="upload_tbl" border="1" style="margin: auto">
        <tr>
            <td>Select File:</td>
            <td><asp:FileUpload ID="fileupload1" CssClass="custom_text" runat="server" /></td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: right"><asp:Button ID="btnUpload" Text="Upload" runat="server"/>&nbsp<asp:Button ID="btnCancel" Text="Cancel" runat="server"/></td>
        </tr>
    </table>
            

        </div>
        <asp:DataGrid ID="uploadgrid" Visible="false" AutoGenerateColumns="true" runat="server"></asp:DataGrid>
    </div>

</asp:Content>
