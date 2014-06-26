using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using TestModelTTT;
using ZKJC.CommonDAO;

namespace ExportXLSFromDataTable
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            #region MyRegion
            //System.IO.Stream strm = ExportXLSFromDataTable.ExportDataTable.DataSource();
            //byte[] bytes = new byte[strm.Length];

            //strm.Read(bytes, 0, bytes.Length);
            //strm.Close();

            //// Response.addheader("content-disposition", "attachment; filename=" );
            ////Response.Charset = "GB2312";
            //////Response.ContentEncoding = System.Text.Encoding.UTF8;
            ////// 添加头信息，为"文件下载/另存为"对话框指定默认文件名 
            //Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlEncode("导出文件.xls"));
            ////// 添加头信息，指定文件大小，让浏览器能够显示下载进度 
            //Response.AddHeader("Content-Length", bytes.Length.ToString());
            ////// 指定返回的是一个不能被客户端读取的流，必须被下载 
            //Response.ContentType = "application/octet-stream";
            //// 把文件流发送到客户端 

            //Response.BinaryWrite(bytes);
            //Response.Flush();
            //Response.End(); 
            #endregion

            SqlServerDAO dao = new SqlServerDAO();
            ModelA test = new ModelA();
            test.DictID = "AttState";
            test = dao.GetEntity<ModelA>(test, "DictID");
            litTest.Text = test.DictName;

        }
    }
}