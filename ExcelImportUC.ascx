<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ExcelImportUC.ascx.cs" Inherits="ExcelImportUC" %>
<div class="container">
    <br/>
    <div class="row">
        <div class="col-sm-8">
            <asp:FileUpload runat="server" ID="flimport" />
        </div>
        <div class="col-sm-4">
            <asp:Button runat="server" ID="btnUpload" CssClass="btn-primary" Text="Import" OnClick="btnUpload_OnClick" style="display: none" />
        </div>
    </div>
   <br />
    <div class="row">
        <div class="col-sm-12">
            <asp:GridView runat="server" CssClass="table table-striped table-bordered table-hover" ID="grdExcelData" AutoGenerateColumns="false"
                OnRowDataBound="grdExcelData_RowDataBound"
                OnRowCancelingEdit="grdExcelData_RowCancelingEdit"
                OnRowUpdating="grdExcelData_RowUpdating"
                OnRowEditing="grdExcelData_RowEditing">
                <Columns >
                    <asp:TemplateField HeaderStyle-CssClass="hidden-xs">
                        <ItemTemplate>
                            <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandName="Edit" />
                        </ItemTemplate>
                        <EditItemTemplate>

                            <asp:Button ID="btnCancel" runat="server" Text="Cancel" CommandName="Cancel" />
                              <asp:Button ID="btnUpdate" runat="server" Text="Validate" CommandName="Update" />
                        </EditItemTemplate>
                        <HeaderStyle CssClass="headerCommandClass"></HeaderStyle>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Name">
                        <ItemTemplate>
                            <asp:Label ID="lbl_Name" runat="server" Text='<%#Eval("Name") %>'></asp:Label>
                        </ItemTemplate>
                          <HeaderStyle CssClass="headerClass"></HeaderStyle>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Grade" HeaderStyle-CssClass="hidden-xs">
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
                          <HeaderStyle CssClass="headerClass"></HeaderStyle>

                    </asp:TemplateField>

                </Columns>

            </asp:GridView>
        </div>
    </div>

    <div class="row">
        <div class="col-sm-12 pull-left">
            <asp:Button ID="btnValidateAndSave" Visible="False" runat="server" Text="Submit" CssClass="btn-primary" OnClick="btnValidateAndSave_Click" />
        </div>
    </div>

</div>
<script type="text/javascript">
    function UploadFile(flimport) {
        if (flimport.value !== '') {
            document.getElementById("<%=btnUpload.ClientID %>").click();
        }
    }
</script>