<%@ Page Title="MR Details" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="MRDetails.aspx.vb" Inherits="TO.MRDetails" EnableEventValidation="false" %>

<asp:Content ID="MRDetailsContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="documentno_holder">
        <%--<asp:Image ID="Image2" CssClass="MRDetails_image" runat="server" ImageUrl="~/images/image1.bmp" />--%>
        <asp:Label ID="DocNoLabel" CssClass="lbl_css" runat="server" Text=""></asp:Label>
        <asp:Button ID="btnlogout" CssClass="btn_css_xs2" runat="server" UseSubmitBehavior="false" Text="Logout" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" />
        <asp:Button ID="buttondelete" CssClass="btn_css_xs1 btn_css_delete" runat="server" UseSubmitBehavior="false" Text="Cancel This Document" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" />
        <asp:Button ID="btn_delete_whs_negative" CssClass="btn_css_xs1 btn_css_delete btn_delete_whs_negative_css" runat="server" UseSubmitBehavior="false" Text="Delete all items w/ negative and zero WHSE QTY" Style="font-family: Verdana, Geneva, Tahoma, sans-serif" />
       <%-- <asp:Button ID="btn_delete_whs_zero" CssClass="btn_css_xs1 btn_css_delete btn_delete_whs_negative_css" runat="server" UseSubmitBehavior="false" Text="Delete all items w/ zero WHSE QTY" Style="font-family: Verdana, Geneva, Tahoma, sans-serif" />--%>
        <asp:Button ID="btn_delete_qtyavail_negative" CssClass="btn_css_xs1 btn_css_delete btn_delete_qtyavail_negative_css" runat="server" Visible="false" UseSubmitBehavior="false" Text="Delete all items w/ negative QTY AVAIL" Style="font-family: Verdana, Geneva, Tahoma, sans-serif" />
        <asp:Button ID="buttonSave" CssClass="btn_css_xs2" runat="server" Text="Save All" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" />
        <asp:Button ID="BtnHome" CssClass="btn_css_xs2" runat="server" UseSubmitBehavior="false" Text="Home" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" />
        <%--            <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="~\images\Save.png" CssClass="btn_css_xs2" ToolTip="Click to Add" />
            <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~\images\Home.png" CssClass="btn_css_xs2" ToolTip="Click to Add" />--%>
    </div>
    <div class="div_controls">
        <asp:Label ID="Label2" runat="server" Text="Main Category : " Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 12px; text-align: center"></asp:Label>
        <asp:DropDownList ID="dropdownMainCat" CssClass="custom_text" runat="server" AutoPostBack="true" Width="200px" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px"></asp:DropDownList>
        &nbsp&nbsp

                    <asp:Label Text="Item Code :" runat="server"></asp:Label>
        <asp:TextBox ID="txt_itemcode" CssClass="custom_text" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" runat="server" Width="150" AutoPostBack="true"></asp:TextBox>&nbsp&nbsp
            <asp:Label ID="Label1" runat="server" Text="Item List : " Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 12px"></asp:Label>
        <asp:DropDownList ID="dropdownItem" runat="server" Width="500px" CssClass="custom_text" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px"></asp:DropDownList>
        &nbsp &nbsp
        <asp:ImageButton ID="addItems" runat="server" ImageUrl="~\images\Add2.png" CssClass="AddItem_css" ToolTip="Click to Add New Item" />
        <asp:ImageButton ID="export_btn" runat="server" ImageUrl="~\images\download3.png" CssClass="AddItem_css" ToolTip="Download to Excel" />
        <asp:ImageButton ID="upload_btn" runat="server"  ImageUrl="~\images\open.png" CssClass="AddItem_css" ToolTip="Upload" />
        <%--            <asp:Button ID="buttonAdd" CssClass="btn_css btn_css2" runat="server" Text="Add Item" width="100px" style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" />--%>
    </div>



    <div class="row MR_div" id="gv">

        <div class="main_holder" id="div_message_box" runat="server">
            <div id="msgboxdiv2" class="msgbox_div2">
                <asp:Label Text="Are you sure you want to delete this document?" ID="msgbox2_label" runat="server" CssClass="msgbox2_label"></asp:Label><p></p>
                <br />
                <asp:Button ID="button1" CssClass="" runat="server" Text="Delete" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" />
                <asp:Button ID="button2" CssClass="" runat="server" Text="Cancel" Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px" />
            </div>
        </div>
        <%--            <div>
        <h2>Merchandise Replenishment - Details</h2>
    </div>--%>

        <%--    <div>
        <p>--%>
        <%--            <asp:Label ID="Label2" runat="server" Text="Main Category : " style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 12px; text-align: center"></asp:Label>
            <asp:DropDownList ID="dropdownMainCat" CssClass="custom_text" runat="server" AutoPostBack="true" width="250px" style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 11px"></asp:DropDownList>
        --%>
        <%--        </p>
    </div>
    
    <hr />--%>

        <%--    <div class="row">
        <p>
            

        </p>
    </div>--%>



        <asp:UpdatePanel ID="gvpanel" runat="server">

            <Triggers>
                <asp:PostBackTrigger ControlID="GridView1" />
            </Triggers>
            <ContentTemplate>


                <asp:Label ID="Label3" runat="server" Text="Label" Visible="false"></asp:Label>

                <%--<div class="mrdetails_div">--%>

                    <asp:GridView ID="GridView1" runat="server" CssClass="custom_grid custom_grid_extra_css custom_grid_test" RowStyle-Wrap="false"  AutoGenerateDeleteButton="true"
                        AutoGenerateColumns="false"
                        OnRowDataBound="gv_RowDataBound"
                        OnRowCreated="gv_RowCreated"
                        OnRowEditing="gv_RowEditing"
                        OnRowDeleting="gv_RowDeleting"
                        OnRowCancelingEdit="gv_CancelEdit"
                        OnRowUpdating="gv_RowUpdating"
                        OnSelectedIndexChanged="gv_SelectedIndexChanged"
                        OnSelectedIndexChanging="gv_SelectedIndexChanging"
                        Style="font-family: Verdana, Geneva, Tahoma, sans-serif; font-size: 10px;">
                        <%--                    <HeaderStyle CssClass="GridviewScrollHeader" /> 
                    <RowStyle CssClass="GridviewScrollItem" />
                    <PagerStyle CssClass="GridviewScrollPager" />--%>

                        
                       <Columns>
                      <asp:TemplateField>
                        <HeaderTemplate>
                         <input id="chkAll" type="checkbox" />
                      </HeaderTemplate>
                      <ItemTemplate>
                      <asp:CheckBox ID="chkSelect" runat="server" />
                      </ItemTemplate>
                      </asp:TemplateField>
                      </Columns>


                    </asp:GridView>

                    <asp:HiddenField ID="hfGridView1SV" runat="server" />
                    <asp:HiddenField ID="hfGridView1SH" runat="server" />

               <%-- </div>--%>


            </ContentTemplate>
        </asp:UpdatePanel>

    </div>
    <script type="text/javascript" src="Scripts/gridviewscroll.js"></script>

<%--    <script type="text/javascript">
 window.onload = function () {
     var gridViewScroll = new GridViewScroll({
         elementID: "GridView1" // Target element id
     });
     gridViewScroll.enhance();
 }
</script>

    <script type="text/javascript">
        var gridViewScroll = new GridViewScroll({
            elementID: "GridView1", // String
            width : 300, // Integer or String(Percentage)
            height : 130, // Integer or String(Percentage)
            freezeColumn : true, // Boolean
            freezeFooter : false, // Boolean
            freezeColumnCssClass: "", // String
            freezeFooterCssClass: "", // String
            freezeHeaderRowCount : 1, // Integer
            freezeColumnCount : 3,// Integer
            onscroll: function (scrollTop, scrollLeft) {}// onscroll event callback
            });
        
    </script>

    <script type="text/javascript">
    var gridViewScroll = new GridViewScroll({
        elementID: "GridView1"
    });
    gridViewScroll.enhance();
    var scrollPosition = gridViewScroll.scrollPosition // get scroll position
    var scrollTop = scrollPosition.scrollTop;
    var scrollLeft = scrollPosition.scrollLeft;

    var scrollPosition = { scrollTop: 50, scrollLeft: 50};
    gridViewScroll.scrollPosition = scrollPosition; // set scroll position
</script>

    <script type="text/javascript">
    var gridViewScroll = new GridViewScroll({
        elementID: "GridView1"
    });
    gridViewScroll.enhance();
    var scrollPosition = gridViewScroll.scrollPosition // get scroll position
    var scrollTop = scrollPosition.scrollTop;
    var scrollLeft = scrollPosition.scrollLeft;

    var scrollPosition = { scrollTop: 50, scrollLeft: 50};
    gridViewScroll.scrollPosition = scrollPosition; // set scroll position
</script>

    <script type="text/javascript">
    var gridViewScroll = new GridViewScroll({
        elementID: "GridView1"
    });
    gridViewScroll.enhance(); // Apply the gridviewscroll features
    gridViewScroll.undo(); // Undo the DOM changes, And remove gridviewscroll features
</script>

    <script type="text/javascript">
    var gridViewScroll = new GridViewScroll({
        elementID: "GridView1",
        onscroll: function (scrollTop, scrollLeft) {
            console.log("scrollTop: " + scrollTop + ", scrollLeft: " + scrollLeft);
        }
    });
    gridViewScroll.enhance();
</script>--%>

    <script src="Scripts/jquery-1.4.1.min.js" type="text/javascript"></script>
    <script src="Scripts/ScrollableGridPlugin.js" type="text/javascript"></script>

    <script type="text/javascript">
        $(document).ready(function () {
            $('#<%=GridView1.ClientID %>').Scrollable({
                ScrollHeight: $(window).height() - 130
        });
    });

    </script>
       <%-- <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>--%>
    <script type="text/javascript">
        $(document).ready(function () {
            var allInputs = $(':text:visible'); //(1)collection of all the inputs I want (not all the inputs on my form)
            $(":text").on("keydown", function () {//(2)When an input field detects a keydown event
                if (event.keyCode == 13) {
                    event.preventDefault();
                    var nextInput = allInputs.get(allInputs.index(this) + 1);//(3)The next input in my collection of all inputs
                    if (nextInput) {
                        nextInput.focus(); //(4)focus that next input if the input is not null
                    }
                }
            });
        });
    </script>


</asp:Content>
