namespace Ideal.Office.Web
{
    using System;
    using System.IO;
    using System.Web;

    public static class ResponseExtensions
    {
        #region 方法

        public static void ExportExcel(this HttpResponse response, MemoryStream stream, string excelName)
        {
            excelName += DateTime.Now.ToString("_yyyyMMddhhmmss");
            excelName = excelName.IndexOf(".xls") == -1 ? (excelName + ".xls") : excelName;
            response.ContentType = "application/x-msdownload";
            response.Charset = "";
            response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(excelName, System.Text.Encoding.UTF8));
            response.BinaryWrite(stream.GetBuffer());
            response.End();
        }

        #endregion
    }
}
