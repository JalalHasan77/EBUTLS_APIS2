<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Test.aspx.vb" Inherits="InvoicingAPI.ViewReport" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="~/css/Forms.css" rel="stylesheet" />
    <style type="text/css">
        .style2 {
            width: 100%;
        }

        .style1 {
            width: 100%;
        }
        .auto-style4 {
            width: 98%;
            height: 44px;
        }
        .auto-style5 {
            width: 2%;
            height: 44px;
        }
        .auto-style6 {
            width: 64px;
        }
        .auto-style7 {
            width: 1%;
            height: 28px;
        }
        .auto-style8 {
            width: 3%;
            height: 28px;
        }
        .auto-style9 {
            width: 20%;
            height: 28px;
        }
        </style>
</head>
<body style="margin: 0px;" class="FormBody">
    <form id="form1" runat="server" class="MeForm">
        <table class="FormHeader" style="width: 100%; ">
            <tr>
                <td style="text-align: left; vertical-align: middle; font-family: Arial; font-size: 10px; " class="auto-style4">
                    <asp:Label ID="lblAppName" runat="server" Font-Bold="False" Font-Size="24px" ForeColor="#333333" Font-Names="Arial Narrow">Invoice</asp:Label>
                </td>
                
                <td style="vertical-align: middle; text-align: right; font-family: Arial; font-size: 12px; " class="auto-style5">
                    <table cellpadding="2" cellspacing="0" class="style2">
                        <tr>
                            <td>
                                <asp:Button ID="cmdSave" runat="server" CssClass="MeBoxButton" Text="Save" Width="60px"  />
                            </td>
                            <td class="auto-style6">
                                <asp:Button ID="cmdClose" runat="server" CssClass="MeBoxButton" Text="Close" Width="60px" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table cellpadding="6" cellspacing="0" class="style3" style="background-color: #FFFFFF; font-family: Arial; font-size: 14px;">
            <tr>
                <td style="width: 2%; vertical-align: top;">
                    <table cellpadding="0" cellspacing="0" class="style2">
                        <tr>
                            <td>
                                <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/159278-200.png" Height="84px" Width="91px" />
                            </td>
                        </tr>
                        <tr>
                            <td style="margin-left: 40px">
                                <asp:Label ID="lbl_Invoice_ID" runat="server" Visible="False"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="margin-left: 40px">
                                <asp:Label ID="lbl_Mode" runat="server" Visible="False"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="margin-left: 40px">
                                <asp:Label ID="lbl_check" runat="server" Visible="False"></asp:Label>
                            </td>
                        </tr>

                    </table>
                </td>
                <td style="width: 98%">
                    <table cellpadding="0" cellspacing="0" class="style2">
                        <tr>
                            <td style="width: 2%; vertical-align: top;">
                                <table cellpadding="6" cellspacing="0" class="style3">
                                    <tr>
                                        <td style="width: 2%; margin-left: 80px;" aria-multiline="False">
                                            <asp:Label ID="Label151" runat="server" Text="Customer No." Width="138px"></asp:Label>
                                        </td>
                                        <td>
                                            <asp:TextBox ID="txt_CustNo" runat="server" Width="514px" CssClass="MeBox"></asp:TextBox>
                                        </td>   
                                    </tr>
                                    
                                    </table>
                            </td>
                        </tr>
                        </table>
                </td>
            </tr>
        </table>
     <asp:Panel ID="Panel1" runat="server">
          
       <table cellpadding="10" cellspacing="0" class="style2">
         <tr>
            <td>
              <asp:MultiView ID="MultiView1" runat="server" ActiveViewIndex="0">
                 <asp:View ID="View1" runat="server">
                     <table cellpadding="4" cellspacing="0" class="style1">
                         <tr>
                             <td style="Width: 2%">
                            </td>
                              <td style="Width: 70%; text-align:center">
                                    <asp:Label ID="Label1" runat="server" Text="Item" ></asp:Label>
                                 </td>
                              <td style="Width: 20%;text-align:center">
                                     <asp:Label ID="Label6" runat="server" Text="Date of Supply" ></asp:Label>
                                 </td>
                                <td style="Width: 30%;text-align:center">
                                     <asp:Label ID="Label2" runat="server" Text="Price" ></asp:Label>
                                 </td>
                             <td>

                             </td>
                         </tr>
                             <tr>
                                <td style="Width: 2%">
                                    <asp:ImageButton ID="cmdAdd_Item" runat="server" CausesValidation="False" ImageUrl="~/Images/Grids/plus-24.png" OnClick="cmdAdd_Item_Click" Text="Add new" ToolTip="Add new Item" />
                                </td>
                                
                                 <td style="Width: 70%; text-align:center;">    
                                     <asp:TextBox ID="txt_Item" runat="server" Width="379px"></asp:TextBox>
                                 </td>
                                 <td style="Width: 20%; text-align:center;">    
                                     <asp:TextBox ID="txt_Date" runat="server" Width="359px"></asp:TextBox>
                                 </td>
                                 
                                 <td style="text-align:center;">
                                     
                                     <asp:TextBox ID="txt_Price" runat="server" Width="237px" CssClass="MeBox"></asp:TextBox></td>
                             </tr>
                            </table>
                   <asp:GridView ID="gvItems" runat="server" OnRowDeleting="cmdDelete_Click" AutoGenerateColumns="False" GridLines="Horizontal" meta:resourcekey="gvApplicantsResource1" Style="font-size: small; text-align: left;" Width="100%" CellPadding="4" >
                     <Columns>
                       <asp:TemplateField ShowHeader="False">
                        
                         <ItemTemplate>
                          <table cellpadding="2" cellspacing="0" class="style1">
                             <tr>
                                 <td style="width:2%">


                                 </td>
                                
                             </tr>
                          </table>
                                            </ItemTemplate>
                                            
                                            <ItemStyle Width="1%" />
                                        </asp:TemplateField>
                                   
                                    <asp:BoundField DataField="ID" HeaderText="ID" />
                                    <asp:BoundField DataField="DATE_OF_SUPPLY" HeaderText="Date_Of_Supply" />
                                    <asp:BoundField DataField="DESCRIPTION" SortExpression="Item" HeaderText="Description" />
                                    <asp:BoundField DataField="DUE_AMOUNT" SortExpression="Price" HeaderText="Due_Amount" />
                                     <asp:CommandField ShowDeleteButton="True" ButtonType="Button" />
                                    </Columns>
                                    <EmptyDataTemplate>
                                    
                                    </EmptyDataTemplate>
                                    <HeaderStyle HorizontalAlign="Left" />

                                    <RowStyle Height="30px" />

                                </asp:GridView>
                            </asp:View>
            
                            </asp:MultiView>
                        &nbsp;</td>
                </tr>
          
            </table >
           <asp:Panel ID="Panel2" runat="server" GroupingText="Items Total" Width="10%">
                                                    <table cellpadding="10" cellspacing="0" class="style2">
                                                        <tr>
                                                            <td>
                                                                <asp:Label ID="lblTotal" runat="server"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </asp:Panel>
            <br />

         <table cellpadding="2" cellspacing="2" class="style1">
             <tr>
                 <td class="auto-style7">
                     &nbsp
                 </td>
                 <td class="auto-style8">
                     <asp:Label ID="Label4" runat="server" Text="Transaction Fees"></asp:Label>
                 </td>
                 <td class="auto-style9">
                     <asp:TextBox ID="txt_Fees" runat="server" Width="20%" AutoPostBack="True"></asp:TextBox>
                 </td>
             </tr>
             <tr>
                 <td class="auto-style7">
                     &nbsp
                 </td>
                 <td class="auto-style8">
                     <asp:Label ID="Label5" runat="server" Text="Vat"></asp:Label>
                 </td>
                 <td class="auto-style9">
                     <asp:TextBox ID="txt_Vat" runat="server" Width="20%" AutoPostBack="True"></asp:TextBox>
                 </td>
             </tr>
             <tr>
                 <td class="auto-style7">
                     &nbsp
                 </td>
                 <td class="auto-style8">
                     <asp:Label ID="Label3" runat="server" Text="Discount"></asp:Label>
                 </td>
                 <td class="auto-style9">
                     <asp:TextBox ID="txt_Discount" runat="server" Width="20%" CssClass="MeBox" AutoPostBack="True"></asp:TextBox>
                 </td>
             </tr>
           
         </table>
       
        </asp:Panel>
         <asp:Panel ID="Panel3" runat="server" GroupingText="Total" Width="10%">
                                                    <table cellpadding="10" cellspacing="0" class="style2">
                                                        <tr>
                                                            <td>
                                                                <asp:Label ID="lblTo" runat="server"></asp:Label>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </asp:Panel>
        
    </form>
</body>
</html>
