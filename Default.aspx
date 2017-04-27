<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>
<%@ Register src="ExcelImportUC.ascx" tagPrefix="excimp" tagName="imp"%>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
      <script src="Scripts/jquery-1.9.1.min.js"></script>
    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <script src="Scripts/bootstrap.min.js"></script>
    <link href="Content/custom.css" rel="stylesheet" />
    <script src="Scripts/bootstrap-select.min.js"></script>
    <link href="Content/bootstrap-select.min.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <excimp:imp ID="ct1" runat="server" />
    </div>
    </form>
</body>
 
</html>
