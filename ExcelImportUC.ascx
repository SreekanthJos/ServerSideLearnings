<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ExcelImportUC.ascx.cs" Inherits="ExcelImportUC" %>
<div class="container">
    <br/>
    <div class="row">
        <div class="col-sm-8">
            <asp:FileUpload runat="server" ID="flimport" />
        </div>
        <div class="col-sm-4">
            <asp:Button runat="server" ID="btnUpload" CssClass="btn-primary" Text="Import" OnClick="btnUpload_OnClick" />
        </div>
    </div>
   <br />
    <div class="row">
        <div class="col-sm-12">
            <asp:GridView runat="server" CssClass="table table-striped table-bordered table-hover" ID="grdExcelData" AutoGenerateColumns="false"
                OnRowDataBound="grdExcelData_RowDataBound"
                OnRowCancelingEdit="grdExcelData_RowCancelingEdit"
                OnRowEditing="grdExcelData_RowEditing">
                <Columns >
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandName="Edit" />
                        </ItemTemplate>
                        <EditItemTemplate>

                            <asp:Button ID="btnCancel" runat="server" Text="Cancel" CommandName="Cancel" />
                        </EditItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Name">
                        <ItemTemplate>
                            <asp:Label ID="lbl_Name" runat="server" Text='<%#Eval("Name") %>'></asp:Label>
                        </ItemTemplate>

                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Grade">
                        <ItemTemplate>
                            <asp:Label ID="lblGrade" runat="server" Text='<%#Eval("grade") %>'></asp:Label>
                        </ItemTemplate>
                        <EditItemTemplate>
                            <asp:DropDownList ID="ddlGrades" CssClass="selectpicker" runat="server"></asp:DropDownList>
                        </EditItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Rank">
                        <ItemTemplate>
                            <asp:Label ID="lblRank" runat="server" Text='<%#Eval("Rank") %>'></asp:Label>
                        </ItemTemplate>
                        <EditItemTemplate>

                            <asp:DropDownList ID="ddlRanks" CssClass="selectpicker" runat="server"></asp:DropDownList>

                        </EditItemTemplate>
                    </asp:TemplateField>

                </Columns>

            </asp:GridView>
        </div>
    </div>

    <div class="row">
        <div class="col-sm-12 pull-left">
            <asp:Button ID="btnValidateAndSave" Visible="False" runat="server" Text="ReValidate" CssClass="btn-primary" OnClick="btnValidateAndSave_Click" />
        </div>
    </div>

</div>
