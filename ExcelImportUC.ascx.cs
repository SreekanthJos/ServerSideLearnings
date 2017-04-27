using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.UI.WebControls;

public partial class ExcelImportUC : System.Web.UI.UserControl
{
    private List<string> grades = null;
    private List<int> ranks = null;
    private List<int> rowIndices = null;
    protected void Page_Load(object sender, EventArgs e)
    {
        grades = new List<string>() { "First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Nineth", "Tenth" };
        ranks = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
        rowIndices = new List<int>();
    }

    protected void btnUpload_OnClick(object sender, EventArgs e)
    {
        if (flimport.HasFile)
        {
            string fileName = Path.GetFileName(flimport.PostedFile.FileName);
            string folderPath = ConfigurationManager.AppSettings["folderPath"];
            Random rndm = new Random();
            string filePath = Server.MapPath(folderPath + "\\" + rndm.Next() + fileName);
            flimport.SaveAs(filePath);
            ImportDataToGrid(filePath);
        }
    }

    private void ImportDataToGrid(string filePath)
    {
        Application excelApp = null;
        Workbook workbook = null;
        try
        {
            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(filePath);
            var worksheet = workbook.Worksheets[1] as
               Microsoft.Office.Interop.Excel.Worksheet;
            Range excelRange = worksheet.UsedRange;

            worksheet = null; // now No need of this so should expire.

            //Reading Excel file.
            object[,] valueArray = (object[,])excelRange.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);

            excelRange = null; // you don't need to do any more Interop. Now No need of this so should expire.

            System.Data.DataTable dt = ProcessObjects(valueArray);
            ViewState["data"] = dt;
            BindDatatoGrid(dt);
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            #region Clean Up

            if (workbook != null)
            {
                #region Clean Up Close the workbook and release all the memory.

                workbook.Close(false, filePath, Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

                #endregion Clean Up Close the workbook and release all the memory.
            }
            workbook = null;

            if (excelApp != null)
            {
                excelApp.Quit();
            }
            excelApp = null;

            #endregion Clean Up
        }
    }

    private void BindDatatoGrid(System.Data.DataTable dt)
    {
        grdExcelData.DataSource = dt;
        grdExcelData.DataBind();
    }

    private System.Data.DataTable ProcessObjects(object[,] valueArray)
    {
        System.Data.DataTable dt = new System.Data.DataTable();

        #region Get the COLUMN names

        for (int k = 1; k <= valueArray.GetLength(1); k++)
        {
            dt.Columns.Add((string)valueArray[1, k]);  //add columns to the data table.
        }

        #endregion Get the COLUMN names

        #region Load Excel SHEET DATA into data table

        object[] singleDValue = new object[valueArray.GetLength(1)];
        //value array first row contains column names. so loop starts from 2 instead of 1
        for (int i = 2; i <= valueArray.GetLength(0); i++)
        {
            for (int j = 0; j < valueArray.GetLength(1); j++)
            {
                if (valueArray[i, j + 1] != null)
                {
                    singleDValue[j] = valueArray[i, j + 1].ToString();
                }
                else
                {
                    singleDValue[j] = valueArray[i, j + 1];
                }
            }
            dt.LoadDataRow(singleDValue, System.Data.LoadOption.PreserveChanges);
        }

        #endregion Load Excel SHEET DATA into data table

        return (dt);
    }

    protected void grdExcelData_RowDataBound(object sender, GridViewRowEventArgs e)
    {

        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            DataRowView dr = e.Row.DataItem as DataRowView;

            if ((e.Row.RowState & DataControlRowState.Edit) > 0)
            {

                DropDownList ddlGrades = (DropDownList)e.Row.FindControl("ddlGrades");
                ddlGrades.DataSource = grades;
                ddlGrades.DataBind();
                ddlGrades.SelectedValue = dr["grade"].ToString();

                DropDownList ddlRanks = (DropDownList)e.Row.FindControl("ddlRanks");
                ddlRanks.DataSource = ranks;
                ddlRanks.DataBind();
                ddlRanks.SelectedValue = dr["rank"].ToString();
            }
            else
            {
                System.Web.UI.WebControls.Label lbl = (System.Web.UI.WebControls.Label)e.Row.FindControl("lblgrade");

                bool isExist = grades.Contains(lbl.Text);

                if (!isExist)
                {
                    e.Row.Cells[2].BackColor = System.Drawing.Color.Red;
                    btnValidateAndSave.Visible = true;
                }
            }



        }
    }

    protected void grdExcelData_RowEditing(object sender, GridViewEditEventArgs e)
    {
        grdExcelData.EditIndex = e.NewEditIndex;
        rowIndices.Add(e.NewEditIndex);
        ViewState["RowIndices"] = rowIndices;
        if (ViewState["data"] != null)
        {
            System.Data.DataTable dt = (ViewState["data"] as System.Data.DataTable);
            BindDatatoGrid(dt);
        }
    }

    protected void grdExcelData_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
    {
        grdExcelData.EditIndex = -1;
        rowIndices.Remove(grdExcelData.EditIndex);
        ViewState["RowIndices"] = rowIndices;
        if (ViewState["data"] != null)
        {
            System.Data.DataTable dt = (ViewState["data"] as System.Data.DataTable);
            BindDatatoGrid(dt);
        }
    }

    protected void btnValidateAndSave_Click(object sender, EventArgs e)
    {
        System.Data.DataTable dt = null;
        if (ViewState["RowIndices"] != null)
        {
           
            rowIndices = ViewState["RowIndices"] as List<int>;
            foreach (var index in rowIndices)
            {
                GridViewRow row = grdExcelData.Rows[index];
                string grade = ((DropDownList)row.Cells[2].FindControl("ddlGrades")).SelectedValue;
                string rank = ((DropDownList)row.Cells[2].FindControl("ddlRanks")).SelectedValue;

                if (ViewState["data"] != null)
                {
                    dt = (ViewState["data"] as System.Data.DataTable);
                    dt.Rows[index][1] = grade;
                    dt.Rows[index][2] = rank;

                    
                }
            }
            if (dt!=null)
            {
                ViewState["data"] = dt;
                grdExcelData.EditIndex = -1;
                BindDatatoGrid(dt);
            }
        }
    }
}