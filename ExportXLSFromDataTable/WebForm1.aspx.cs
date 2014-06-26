using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExportXLSFromDataTable
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebForm1_LoadComplete;
            this.InitLoading();
            Thread.Sleep(5000);
            Response.Write("哈哈");
        }

        protected void WebForm1_LoadComplete(object sender, EventArgs e)
        {
            this.UnloadLoading();
        }

        /// <summary>
        /// 初始化页面加载效果
        /// </summary>
        public void InitLoading()
        {
            Response.Write("<style>");
            Response.Write("#loader_container {text-align:center; position:absolute; top:40%; width:100%; left: 0;}");
            Response.Write("#loader {font-family:Tahoma, Helvetica, sans; font-size:11.5px; color:#000000; background-color:#FFFFFF; padding:10px 0 16px 0; margin:0 auto; display:block; width:130px; border:1px solid #5a667b; text-align:left; z-index:2;}");
            Response.Write("</style>");
            Response.Write("<div id=loader_container>");
            Response.Write("<div id=loader>");
            Response.Write("<div align=center><img src='http://jimpunk.net/Loading/wp-content/uploads/loading2.gif'/></div>");
            Response.Write("</div>");
            Response.Write("</div></div>");
            Response.Flush();
        }
        /// <summary>
        /// 加载完毕
        /// </summary>
        public void UnloadLoading()
        {
            Response.Write(" <script language=JavaScript type=text/javascript>");
            Response.Write("var targelem = document.getElementById('loader_container');");
            Response.Write("targelem.style.display='none';");
            Response.Write("targelem.style.visibility='hidden';");
            Response.Write(" </script>");
        }
    }
}