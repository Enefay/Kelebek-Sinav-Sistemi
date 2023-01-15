<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="KelebekSistemi.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css" crossorigin="anonymous" />
    <style>
        .LblHata {
            color: red;
            font-weight: bold;
            font-size: 2em;
            font-family: Arial, Helvetica, sans-serif;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:FileUpload ID="fluExcell" runat="server"></asp:FileUpload>
            <asp:Button ID="btnUpload" runat="server" Text="Yükle" OnClick="btnUpload_Click" />
            <asp:Label ID="Lblmessage" runat="server"></asp:Label>
            <asp:Label class="LblHata" ID="Lblhata" runat="server"></asp:Label>
        </div>
        <div>
            <asp:GridView ID="grdExcell" runat="server"></asp:GridView>
            
        </div>
        <hr />

        <div class="row">
            <div class="col-md-6">
                <table>
                    <asp:Repeater ID="rptExcell" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td><%# Eval("F1") %></td>
                                <td><%# Eval("F2") %></td>
                                <td><%# Eval("F3") %></td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
            </div>
            <div class="col-md-6">
                <div class="row">
                    <div class="col-md-6">
                        <asp:Label class="text-danger text-center" ID="ToplamOgrenciSayisi" runat="server"></asp:Label>
                        <asp:CheckBoxList ID="CheckBoxList" AutoPostBack="true" OnSelectedIndexChanged="CheckBoxList_SelectedIndexChanged" runat="server"></asp:CheckBoxList>
                        <asp:Label ID="SecilenSinif" runat="server"></asp:Label>
                    </div>
                    <div class="col-md-6">
                        <asp:Button ID="KaydetDevamEt" runat="server" Text="Rasgele Kaydet ve Devam Et" OnClick="KaydetDevamEt_Click" Visible="False" />

                        <div class="col-md-6">
                        <asp:TextBox  ID="GirisKagidi" class="mt-5" runat="server" Visible="False"></asp:TextBox>

                        <asp:Button ID="Ara" class="mt-2 float-xl-right"  runat="server" Text="Ara" OnClick="Ara_Click" Visible="False"/>
                    </div>
                    </div>
                </div>
            </div>

        </div>
    </form>
</body>
</html>

