using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using SSharing.Frameworks.Document;
using System.Data;

namespace NpoiDemo.Site.Controllers
{
    public class HomeController : Controller
    {
        private readonly IDocument document = new NPOIDocument();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <returns></returns>
        public ActionResult Export()
        {
            var dt = CreateDataTable();
            var headers = CreateNpoiHeader();

            document.ExportExcelAttachment(dt, heads: headers);

            return Content("OK");
        }

        #region Private method 

        private static DataTable CreateDataTable()
        {
            DataTable dt = new DataTable("测试");

            dt.Columns.Add(new DataColumn("Name", typeof(string)));
            dt.Columns.Add(new DataColumn("Date", typeof(string)));
            dt.Columns.Add(new DataColumn("TotalCount", typeof(int)));
            dt.Columns.Add(new DataColumn("AvgDiscount", typeof(string)));
            dt.Columns.Add(new DataColumn("AvgDiscountRate", typeof(string)));

            for (int i = 0; i < 30000; i++)
            {
                dt.Rows.Add(new object[] { "总部1", "2019-01-01", 10, 10, 12.5 });
                dt.Rows.Add(new object[] { "总部2", "2019-01-01", 10, 10, 12.5 });
                dt.Rows.Add(new object[] { "总部3", "2019-01-01", 10, 10, 12.5 });
            }

            return dt;
        }

        private static List<NpoiHead> CreateNpoiHeader()
        {
            List<NpoiHead> heads = new List<NpoiHead>();
            heads.Add(new NpoiHead("Name", "成本中心", 15));
            heads.Add(new NpoiHead("Date", "日期", 15));
            NpoiHead hc1 = new NpoiHead("", "当天", 15);
            hc1.Childs.Add(new NpoiHead("TotalCount", "张数"));
            hc1.Childs.Add(new NpoiHead("AvgDiscount", "平均折扣"));
            hc1.Childs.Add(new NpoiHead("AvgDiscountRate", "平均折扣率"));
            heads.Add(hc1);

            return heads;
        }
        #endregion
    }
}