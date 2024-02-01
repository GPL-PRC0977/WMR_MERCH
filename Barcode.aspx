<%@ Page Title="Barcode Types" Language="vb" AutoEventWireup="true" MasterPageFile="~/Site.Master" CodeBehind="Barcode.aspx.vb" Inherits="TO.Barcode" %>

<asp:Content ID="BarcodeContent" ContentPlaceHolderID="MainContent" runat="server">

    <h2>Barcode Types</h2>
    <hr />

    <div class="row">
            
        <asp:GridView ID="gvBarcode" runat="server" 
            CssClass="table table-striped" 
            AutoGenerateColumns = "false">
            <%--OnRowDataBound = "OnRowDataBound"
            OnRowEditing = "gv_RowEditing" OnRowCancelingEdit = "gv_CancelEdit"
            OnRowUpdating = "gv_RowUpdating"
            >--%>
        </asp:GridView>
        
    </div>

</asp:Content>
