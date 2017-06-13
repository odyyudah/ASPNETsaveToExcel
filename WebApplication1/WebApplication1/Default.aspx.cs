using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication1
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                dsNorthwind ds = new dsNorthwind();
                dsNorthwindTableAdapters.NorthwindTableAdapter da = new dsNorthwindTableAdapters.NorthwindTableAdapter();
                da.Fill(ds.Northwind);
                var rptDataSource = new ReportDataSource("dsNorthwind", ds.Northwind.AsEnumerable());
                rptDataSource.DataSourceId = "ObjectDataSource1";
                rptDataSource.Name = "dsNorthwind";
                ReportViewer1.LocalReport.DataSources.Add(rptDataSource);
                ReportViewer1.DataBind();
                ReportViewer1.LocalReport.Refresh();
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string extension;
            string filename;

            byte[] bytes = ReportViewer1.LocalReport.Render(
               "EXCELOPENXML", null, out mimeType, out encoding,
                out extension,
               out streamids, out warnings);

            filename = string.Format("{0}.{1}", "Northwind", "xlsx");
            Response.ClearHeaders();
            Response.Clear();
            Response.AddHeader("Content-Disposition", "attachment;filename=" + filename);
            Response.ContentType = mimeType;
            Response.BinaryWrite(bytes);
            Response.Flush();
            Response.End();
        }
    }
}