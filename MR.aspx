<%@ Page Title="Merchandise Replenishment" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="MR.aspx.vb" Inherits="TO.MR" EnableEventValidation="false" %>

<asp:Content ID="MRContent" ContentPlaceHolderID="MainContent" runat="server">
    
    <script type = "text/javascript">

        function CreateFile() {
            var confirm_value = document.createElement("INPUT");
            confirm_value.type = "hidden";
            confirm_value.name = "confirm_value";
            if (confirm("Are you sure you want to create NAV text file?")) {
                return true;
            } else {
                return false;
            }    
        }


        function ShowMsg() {
            alert('There are lines with insufficient quantity.');
        }       

    </script>

    <div class="MR_nav">
        <asp:Image ID="Image1" CssClass="login_imageMR" runat="server" ImageUrl="~/images/Properties.png" />
        <asp:Label ID="sitename" CssClass="site_name" runat="server" Text="Transfer Orders"></asp:Label>

        <asp:Button ID="logout_btn" Text="Logout" runat="server" CssClass="logoutbtn_css" />
        <%--<asp:Button ID="reportsBtn" Text="Generate Reports" runat="server" PostBackUrl="~/Reports.aspx" CssClass="homebtn_css" />--%>
        <asp:Button ID="homebtn" Text="Home" runat="server" CssClass="homebtn_css" />
        
    </div>
    <div class="MR_nav2">
        <asp:Label ID="activeUser" runat="server" ForeColor="gray"></asp:Label>

        <asp:Label ID="lbldate" CssClass="lbldate_css" runat="server"><%= DateTime.Now %></asp:Label>

        
    </div>

<%--            <div class="MR_items">

        </div>--%>
    
    <div class="row MR_div" id="div_filter" runat="server">
        <div class="msgbox_div" id="lbl_msgbox" runat="server">Warning! Document is currently being used by</div>
        
<div class="filter_holder">
    <div class="inside_div">

    </div>
    
    <h4 style="color: white;"><b>Merchandise Replenishment List</b></h4>
    <hr />
            <asp:Label ID="Label2" runat="server" Text="Filter :" style="font-family: Verdana, Geneva, Tahoma, sans-serif; margin-top: 6px; font-size: 12px;" Font-Size="12pt"></asp:Label>
            <asp:DropDownList ID="ddlFields" Autopostback="true" CssClass="custom_text status_dd_css" runat="server" Font-Size="10pt" Width="150px">
                <asp:ListItem>-</asp:ListItem>
                <asp:ListItem>Company</asp:ListItem>
                <asp:ListItem>Brand</asp:ListItem>
                <asp:ListItem>Status</asp:ListItem>
    </asp:DropDownList>
    <asp:DropDownList ID="dd_status" runat="server" CssClass="custom_text status2_dd_css" Width="150px" AutoPostBack="true"></asp:DropDownList>
    <asp:ImageButton ID="reportsBTN2" runat="server" ImageUrl="~\images\report4.png" CssClass="login_imageMR2_22" PostBackUrl="~/replenishment_report.aspx" ToolTip="Generate Reports" />
    <asp:ImageButton ID="downloadbtn" runat="server" ImageUrl="~\images\excel.png" CssClass="login_imageMR2_2" ToolTip="Export to Excel" />
    
</div>

        <asp:UpdatePanel ID="gvpanel" runat="server" ChildrenAsTriggers="true">
            <Triggers>
                <asp:PostBackTrigger ControlID="GridView1" />
            </Triggers>
            <ContentTemplate>
                <asp:GridView ID="GridView1" runat="server" CssClass="table custom_grid grid_align" 
                    AutoGenerateColumns = "false"
                    OnRowDataBound = "OnRowDataBound"
                    OnRowEditing = "gv_RowEditing" 
                    OnRowCancelingEdit = "gv_CancelEdit"
                    OnRowUpdating = "gv_RowUpdating"
                    OnRowDeleting = "gv_RowDeleting"
                    OnSelectedIndexChanged = "gv_SelectedIndexChanged"
                    style="font-family: Arial; width: auto; font-size: 11px"
                    >

                </asp:GridView>
            </ContentTemplate>
        </asp:UpdatePanel>

        <asp:Label ID="MessageLabel" runat="server" Text="" ForeColor="Red" Width="150px" Height="20px"></asp:Label>

    </div>
    

</asp:Content>
