using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Office.Web.Mvc.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult EasyExport()
        {
            return new ExportExcelResult("template1", "EasyExport");
        }

        public ActionResult EasyExportParam()
        {
            return new ExportExcelResult("template2", "EasyExportParam"
                , new Dictionary<string, string>() { 
                     { "@DateNow", DateTime.Now.ToString() }
                    ,{ "@Month", DateTime.Now.Month.ToString() }
                }
                , new Dictionary<string, object>() { 
                    { "@ID", "006" }
                });
        }

        public ActionResult ExportSplit() {
            var minMonth = 6;
            var maxMonth = 7;

            var str1 = "";
            var str2 = "";
            var str = "";
            for (int i = minMonth; i <= maxMonth; i++)
            {
                str = string.Format("{0}月营业税", i);
                str1 += string.Format("ISNULL([{0}], 0.00) as '{1}',", i, str);
                str2 += string.Format("[{0}],", i);
            }
            str1 = str1.TrimEnd(',');
            str2 = str2.TrimEnd(',');

            return new ExportExcelResult("template2", "EasyExportParam"
                , new Dictionary<string, string>() { 
                     { "@DateNow", DateTime.Now.ToString() }
                    ,{ "@Month", DateTime.Now.Month.ToString() }
                }
                , new object[] { 
                    str1, str2
                }
                , new Dictionary<string, object>() { 
                    { "@ID", "006" }
                });
        }

    }
}
