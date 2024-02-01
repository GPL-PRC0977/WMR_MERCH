<%@ Page Title="Import Excel File" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Import.aspx.vb" Inherits="TO.Import" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    
    <h2>Import Excel File</h2>
    <hr />

    <div class="row">
        <div class="col-md-6">
            <table>
                <tr>
                    <td>
                        <p>Company : <br />
                        <asp:DropDownList ID="ddlCompany" runat="server" Width="100px" AutoPostBack="True" Height="23px">
                        </asp:DropDownList>
                        </p>
                    </td>
                    <td></td><td></td>
                    <td>
                        <p>Store : <br />
                        <asp:DropDownList ID="ddlStore" runat="server" Width="500px" Height="23px">
                        </asp:DropDownList>
                        </p>
                    </td>
                    <td></td><td></td>
                    <td>
                        <p>Count Date : <br />
                        <asp:TextBox ID="textboxDate" runat="server" Width="150px" Height="23px"></asp:TextBox>                        
                        </p>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>              
            <table>
                <tr>
                    <td>
                        <p>File Location : <br />
                        <asp:FileUpload ID="fileuploadExcel" runat="server" Height="23px" Width="500px" />
                        </p>
                    </td>
                    <td></td><td></td>
                    <td>
                        <p><asp:Button ID="buttonImport" runat="server" Text="Import" Width="100px" />
                        </p>
                    </td>
                </tr>
            </table>
            <p>&nbsp;</p>
            <table>
                <tr>
                    <td>
                        <asp:Label ID="labelMessage" runat="server" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
            <p>&nbsp;</p>   
            
        </div>
    </div>
    
</asp:Content>
