<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ExcelImportUC.ascx.cs" Inherits="ExcelImportUC" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit"%>
<asp:ScriptManager runat="server" ID="scm"></asp:ScriptManager>

<asp:Label runat="server" id="lblMsg"></asp:Label>
 
<div class="container">
    <br/>
 
    <div class="row">
        <div class="col-sm-8"> 
            <asp:FileUpload runat="server" ID="flimport" /> 
           <ajaxToolkit:AsyncFileUpload ID="fileUpload1"  OnClientUploadComplete="uploadComplete" OnClientUploadError="uploadError"
CompleteBackColor="White" Width="350px" runat="server" UploaderStyle="Modern" UploadingBackColor="#CCFFFF"
ThrobberID="imgLoad" OnUploadedComplete="fileUploadComplete" />
        </div>
        <div class="col-sm-4">
          
            <asp:Button runat="server" ID="btnUpload" CssClass="btn-primary" Text="Import" OnClick="btnUpload_OnClick" style="display: none" />
      
                </div>
    </div>
   <br />
    <asp:HiddenField runat="server" id="FaslReason" />
    <div class="row">
        <div class="col-sm-12">
            <asp:GridView runat="server" CssClass="table table-striped table-bordered table-hover" ID="grdExcelData" AutoGenerateColumns="false"
                OnRowDataBound="grdExcelData_RowDataBound"
                OnRowCancelingEdit="grdExcelData_RowCancelingEdit"
                OnRowUpdating="grdExcelData_RowUpdating" AutoGenerateEditButton="true" 
                OnRowEditing="grdExcelData_RowEditing">
                <Columns >
                    <%--<asp:TemplateField HeaderStyle-CssClass="hidden-xs">
                        <ItemTemplate>
                            <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandName="Edit" />
                        </ItemTemplate>
                        <EditItemTemplate>

                            <asp:Button ID="btnCancel" runat="server" Text="Cancel" CommandName="Cancel" />
                              <asp:Button ID="btnUpdate" runat="server" Text="Validate" CommandName="Update" />
                        </EditItemTemplate>
                        <HeaderStyle CssClass="headerCommandClass"></HeaderStyle>
                    </asp:TemplateField>--%>

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
    <asp:Button runat="server" ID="HiddenButton" OnClick="RefreshGridView" style="display:none;" />

</div>
  
<script type="text/javascript">
    function UploadFile(flimport) {
        if (flimport.value !== '') {

            document.getElementById("<%=btnUpload.ClientID %>").click();
        }
    }
    function uploadComplete() {
        document.getElementById("<%=FaslReason.ClientID %>").value = "xyz";
        console.log(document.getElementById("<%=FaslReason.ClientID %>").value);
        __doPostBack("<%= HiddenButton.ClientID %>", "");
document.getElementById('<%=lblMsg.ClientID %>').innerHTML = "File Uploaded Successfully";
}
// This function will execute if file upload fails
function uploadError() {
document.getElementById('<%=lblMsg.ClientID %>').innerHTML = "File upload Failed.";
}
</script>