<%@ Page Language="vb" MasterPageFile="~/Site.Master" EnableEventValidation="false" AutoEventWireup="false" CodeBehind="ExportToExcel.aspx.vb" Inherits="TO.ExportToExcel" %>

<asp:Content ID="MRContent" ContentPlaceHolderID="MainContent" runat="server">
            <style>
                .export_grid{
                    width: 100%;
                    position: absolute;
                    border: solid 1px lightgray;
                }
                .export_grid th{
                    width: 100%;
                }
            </style>
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="true" CssClass="export_grid"></asp:GridView>
</asp:Content>
