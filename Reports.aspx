<%@ Page Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" EnableEventValidation="false" CodeBehind="Reports.aspx.vb" Inherits="TO.Reports" %>

<asp:Content ID="MRDetailsContent" ContentPlaceHolderID="MainContent" runat="server">


        <div class="MR_nav">
        <asp:Image ID="Image1" CssClass="login_imageMR" runat="server" ImageUrl="~/images/Properties.png" />
        <asp:Label ID="sitename" CssClass="site_name" runat="server" Text="Transfer Orders"></asp:Label>

        <asp:Button ID="logout_btn" Text="Logout" runat="server" CssClass="logoutbtn_css" />
        <%--<asp:Button ID="exportbtn" Text="Generate Reports" runat="server" CssClass="homebtn_css" />--%>
        <asp:Button ID="homebtn" Text="Home" runat="server" CssClass="homebtn_css" />
        
    </div>
<%--    <div class="MR_nav2">
        <asp:Label ID="activeUser" runat="server" ForeColor="gray"></asp:Label>

        <asp:Label ID="lbldate" CssClass="lbldate_css" runat="server"><%= DateTime.Now %></asp:Label>

        
    </div>--%>
        <div class="MR_items">
            <table border="1" class="tbl_reports">
                <tr>
                    <td style="font-size: 13px;padding-right: 20px;width: 280px;"><asp:CheckBox Visible="false" ID="chk_selectAll"  runat="server" Font-Bold="false" Text="" AutoPostBack="true" />
                        <asp:Label Text="Select all stores and brands" Visible="false" CssClass="chk_label" runat="server"></asp:Label>
                            <div class="store_header3">
                                Site Concept
                            </div>
                        <%--<asp:TextBox ID="test1" runat="server"></asp:TextBox>--%>
                            <div class="chklist_holder3 custom_text">
                                <asp:Label ID="selectedSiteConcept_holder" Text="" runat="server" CssClass="selectedStores_css3"></asp:Label>
                                <asp:CheckBox ID="selectAllSiteConcept" runat="server" Text="Select All" CssClass="selectAll" AutoPostBack="true" />
                                <asp:CheckBoxList ID="chk_siteConcept" runat="server" CssClass="chklist_css" AutoPostBack="true"></asp:CheckBoxList>
                            </div>
                    </td>
                    <td>Store Name:</td>
                    <td>
                            <div class="store_header">
                                Select Store Names
                            </div>
                            <div id="div1" class="chklist_holder custom_text">
                                <asp:Label ID="selected_holder" Text="" runat="server" CssClass="selectedStores_css"></asp:Label>
                                <asp:CheckBox ID="selectAll" runat="server" Text="Select All" CssClass="selectAll" AutoPostBack="true" />
                                <asp:CheckBoxList ID="chk_list" runat="server" CssClass="chklist_css" AutoPostBack="true"></asp:CheckBoxList>
                            </div>
                    </td>
                    <%--<td><asp:TextBox ID="storeCollections" Width="2000" runat="server"></asp:TextBox></td>--%>
                    
                    <td rowspan="2">Date From:</td>
                    <td rowspan="2"><asp:TextBox ID="date_from" runat="server" CssClass="date_txt" Width="140" TextMode="date"></asp:TextBox>
                    </td>
                    <td rowspan="2">Date To:</td>
                    <td rowspan="2"><asp:TextBox ID="date_to" runat="server" CssClass="date_txt" Width="140" TextMode="date"></asp:TextBox></td>
                    <td rowspan="2"><asp:Button ID="search_btn" runat="server" Text="Search" /></td>
                    <td rowspan="2"><asp:Button ID="Exportbtn" runat="server" Text="Download" /></td>
                </tr>
                <tr>
                    <td style="font-size: 13px;padding-right: 20px;">
                        <asp:CheckBox ID="chk_withZero"  runat="server" Font-Bold="false" Text=""/>
                        <asp:Label Text="Show unserved items only" CssClass="chk_label" runat="server"></asp:Label>
                    </td>
                    <td>Brand:</td>
                    <td>
                                <div class="store_header2">
                                    Select Brand Names
                                </div>
                                <div class="chklist_holder2 custom_text">
                                    <asp:Label ID="selected_holder_brands" Text="" runat="server" CssClass="selectedStores_css"></asp:Label>
                                    <asp:CheckBox ID="SelectAllBrand" runat="server" Text="Select All" CssClass="selectAll" AutoPostBack="true" />
                                    <asp:CheckBoxList ID="chk_brand" runat="server" CssClass="chklist_css" AutoPostBack="true"></asp:CheckBoxList>
                                </div>

                    </td>

                    <%--<td><asp:TextBox ID="brandsCollections" Width="200" runat="server"></asp:TextBox></td>--%>

                </tr>

            </table>
        </div>

    <div class="background_div" id="backgrounddiv" runat="server">
            <div Class="deletedItems_css" id="deleted_items_div" runat="server">
<div class="item_holder">
            Deleted Items for Document No: &nbsp <asp:Label ID="selectedDocument" runat="server"></asp:Label><p>
            <asp:Button ID="btn_close" Text="Close" runat="server" CssClass="btn_close_css" />

</div>
                <asp:Label ID="datacounter" Text="0" runat="server" CssClass="datacounter_css"></asp:Label>
       <asp:GridView ID="deletedItems" runat="server" CssClass="deletedItems_grid" AutoGenerateColumns="true"></asp:GridView>
    </div>
    </div>

    <div class="main_data_holder">
                <asp:GridView ID="dg_result" runat="server" CssClass="dg_reports_css" HeaderStyle-Wrap="false" RowStyle-Wrap="false" AutoGenerateColumns="true"
                                onrowdatabound ="OnRowDatabound"
                                OnRowEditing ="dg_result_RowEditing"
                                OnSelectedIndexChanged = "dg_result_SelectedIndexChanged"/>
    </div>

    
</asp:Content>
