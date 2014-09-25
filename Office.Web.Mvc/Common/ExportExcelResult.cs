using Ideal.Office.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace Office.Web.Mvc
{
    public class ExportExcelResult : ActionResult
    {
        private string excelName;
        private MemoryStream stream;

        public ExportExcelResult(string templateName, string fileName)
            : this(templateName, fileName, null, null, new object[0])
        {
        }

        public ExportExcelResult(string templateName, string fileName, Dictionary<string, string> consts)
            : this(templateName, fileName, consts, null, new object[0])
        {
        }

        public ExportExcelResult(string templateName, string fileName, params object[] param)
            : this(templateName, fileName, null, null, param)
        {
        }

        public ExportExcelResult(string templateName, string fileName, Dictionary<string, string> consts, params object[] param)
            : this(templateName, fileName, consts, null, param)
        {
        }

        public ExportExcelResult(string templateName, string fileName, Dictionary<string, string> consts = null, object[] dynamicCol = null, params object[] param)
        {
            fileName += DateTime.Now.ToString("_yyyyMMddhhmmss");
            this.excelName = fileName.IndexOf(".xls") == -1 ? (fileName + ".xls") : fileName;

            this.stream = Export.Down(templateName, consts, dynamicCol, param);
        }


        public override void ExecuteResult(ControllerContext context)
        {
            context.RequestContext.HttpContext.Response.AppendHeader("content-disposition", "attachment;filename=" + HttpUtility.UrlEncode(excelName, System.Text.Encoding.UTF8));
            //context.RequestContext.HttpContext.Response.ContentEncoding = Encoding.GetEncoding("gb2312");
            context.RequestContext.HttpContext.Response.ContentType = "application/excel";
            context.RequestContext.HttpContext.Response.BinaryWrite(stream.GetBuffer());
            context.RequestContext.HttpContext.Response.End();
        }
    }
}