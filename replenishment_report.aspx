<%@ Page Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="replenishment_report.aspx.vb" Inherits="TO.replenishment_report" %>

<asp:Content ID="MRContent" ContentPlaceHolderID="MainContent" runat="server">
        <script src="Content/jquery.min.js"></script>

       <script src="Content/jquery.sumoselect.min.js"></script>
           <link href="Content/sumoselect.css" rel="stylesheet">
       <script type="text/javascript">
        $(document).ready(function () {
                   $(<%=dd_brands.ClientID%>).SumoSelect({ selectAll: true });
               });
   </script>


    <div class="replenishment_report_holder">
        <div class="report_filter">
            <table border="1" class="tbl_replenishment_report_filter">
                <tr>
                    <td>From<br /><asp:TextBox TextMode="Date" CssClass="form-control" runat="server" ID="txt_date_from"></asp:TextBox></td>
                    <td>To<br /><asp:TextBox TextMode="Date" CssClass="form-control" runat="server" ID="txt_date_to"></asp:TextBox></td>
                    <%--<td>Company<br /><asp:ListBox ID="dd_company" SelectionMode="Multiple" CssClass="dd_company_css" runat="server"></asp:ListBox></td>--%>
                    <td>Company<br /><asp:DropDownList ID="ddcompany" runat="server" CssClass="form-control" AutoPostBack="true"></asp:DropDownList></td>
                    <td>Brands<br /><asp:ListBox ID="dd_brands" SelectionMode="Multiple" CssClass="form-control dd_css" runat="server"></asp:ListBox></td>
                    <td><asp:Button ID="btn_search" runat="server" CssClass="btn_search_css" Text="Search" /></td>
                    <td><asp:Button ID="btn_download" runat="server" CssClass= "btn_download_css" Text="Download" /></td>
                </tr>
            </table>
            <asp:GridView runat="server" CssClass="dg" ID="dg"></asp:GridView>
        </div>
    </div>
</asp:Content>
