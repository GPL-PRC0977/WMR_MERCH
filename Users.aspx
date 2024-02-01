<%@ Page Title="Users" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Users.aspx.vb" Inherits="TO.Users" %>

<asp:Content ID="UserContent" ContentPlaceHolderID="MainContent" runat="server">

    <h2>Users</h2>
    <hr />

    <div class="row">
            
        <asp:GridView ID="gvUser" runat="server" 
            CssClass="table table-striped"
            AutoGenerateColumns="false"
            AllowPaging="true" OnPageIndexChanging="gvUser_PageIndexChanging" PageSize="10">
            <%--OnRowDataBound = "OnRowDataBound"
            OnRowEditing = "gv_RowEditing" OnRowCancelingEdit = "gv_CancelEdit"
            OnRowUpdating = "gv_RowUpdating"
            >--%>
        </asp:GridView>
        
    </div>

</asp:Content>
