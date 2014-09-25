using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Ideal.Office.Web
{
    using Ideal.Office.Web;
    using System.Data.SqlClient;

    public partial class _Default : Page
    {
        protected String excelContent;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                #region Step1
                var minMonth = 6;
                var maxMonth = 7;

                var str1 = "";
                var str2 = "";
                var str = "";
                List<string> autoCol = new List<string>();
                List<int> cols = new List<int>();
                for (int i = minMonth; i <= maxMonth; i++)
                {
                    cols.Add(i);
                    str = string.Format("{0}月营业税", i);
                    autoCol.Add(str);
                    str1 += string.Format("ISNULL([{0}], 0.00) as '{1}',", i, str);
                    str2 += string.Format("[{0}],", i);
                }
                str1 = str1.TrimEnd(',');
                str2 = str2.TrimEnd(',');


                var stream = Export.Down("template2", null
                 , new object[] { 
                        str1,str2
                    }
                , new Dictionary<string, object>(){
                      { "@EID", 2 }
                    , { "@Year", 2014 }
                    , { "@MinMonth", minMonth }
                    , { "@MaxMonth", maxMonth }
                    , { "@TaxID", 1 }
                });

                Response.ExportExcel(stream, "动态列的操作");

                #endregion

                #region Step2

                //var stream = Export.Down("template1"
                //    , new Dictionary<string, string>(){
                //    { "@DateNow", DateTime.Now.ToString() }
                //    ,{ "@Month", "全能" }
                //}, new Dictionary<string, object>() { 
                //    { "@p1", 73 },
                //    { "@p2", 8 }
                //}); 
                //Response.ExportExcel(stream, "奉贤企业信息管理");

                #endregion

                #region Step3

                //var stream = Export.Down("template3", consts: new Dictionary<string, string>(){
                //    { "@DateNow", DateTime.Now.ToString() }
                //    ,{ "@Month", "Test" }});

                //Response.ExportExcel(stream, "Test");

                #endregion

                //Import.GetImportConfig();

            }
        }

        protected void btnExcelToTable_Click(object sender, EventArgs e)
        {
            var param = new object[1];
            param[0] = 73;
            excelContent = Export.HtmlTable("template1", "奉贤企业信息管理", param); 
        }
    }
}