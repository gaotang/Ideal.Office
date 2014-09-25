namespace Ideal.Office.Web
{
    using Ideal.Office.Data;
    using Ideal.Office.Excel;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Web;
    public partial class Export
    {
        public static string JsonString(string templateName, string excelName = "", object[] dynamicCol = null, int titleIndex = 0, params object[] parameters)
        {
            ExcelEntity entity;
            if (Import.GetXmlConfig().TryGetValue(templateName, out entity))
            {
                var name = excelName + DateTime.Now.ToString("(yyyy年MM月dd日hh时mm分ss秒)") + (entity.Type == ExcelType.Excel2003 ? ".xls" : ".xlsx");
                using (FileStream fs = new FileStream(entity.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    Excel = FillExcel(fs, entity, dynamicCol, parameters);
                    return LoadJson(titleIndex);
                }
            }
            else
            {
                return "";
            }
        }

        public static string JsonString(string templateName, string excelName = "", int titleIndex = 0, params object[] parameters)
        {
            return JsonString(templateName, excelName, null, titleIndex, parameters);
        }

        public static string JsonString(string templateName, string excelName = "", params object[] parameters)
        {
            return JsonString(templateName, excelName, null, 0, parameters);
        }

        private static string LoadJson(int titleIndex = 0)
        {
            var sht = Excel.Sheet;
            //取行Excel的最大行数
            int rowsCount = sht.PhysicalNumberOfRows;
            //为保证Table布局与Excel一样，这里应该取所有行中的最大列数（需要遍历整个Sheet）。
            //为少一交全Excel遍历，提高性能，我们可以人为把第0行的列数调整至所有行中的最大列数。
            int colsCount = sht.GetRow(0).PhysicalNumberOfCells;

            int colSpan;
            int rowSpan;
            bool isByRowMerged;
            var tempTd = "";

            StringBuilder table = new StringBuilder(rowsCount * 32);

            table.Append("[");
            for (int rowIndex = titleIndex; rowIndex < rowsCount; rowIndex++)
            {
                //table.Append("{");
                //tempTd = "{";
                tempTd = "";
                for (int colIndex = 0; colIndex < colsCount; colIndex++)
                {
                    GetTdMergedInfo(rowIndex, colIndex, out colSpan, out rowSpan, out isByRowMerged);
                    //如果已经被行合并包含进去了就不输出TD了。
                    //注意被合并的行或列不输出的处理方式不一样，见下面一处的注释说明了列合并后不输出TD的处理方式。
                    if (isByRowMerged)
                    {
                        continue;
                    }

                    //table.Append("<td");
                    //if (colSpan > 1)
                    //    table.Append(string.Format(" colSpan={0}", colSpan));
                    //if (rowSpan > 1)
                    //    table.Append(string.Format(" rowSpan={0}", rowSpan));
                    //table.Append(">");

                    //table.Append(sht.GetRow(rowIndex).GetCell(colIndex));

                    tempTd += sht.GetRow(rowIndex).GetCell(colIndex);
                    //列被合并之后此行将少输出colSpan-1个TD。
                    if (colSpan > 1)
                        colIndex += colSpan - 1;

                    if (colSpan > 1)
                        colIndex += colSpan - 1;

                    //table.Append(",");
                    tempTd += ",";

                }
                var b = tempTd.Replace(",", "");
                if (!string.IsNullOrEmpty(b))
                {
                    table.Append("{");
                    table.Append(tempTd.Trim(','));
                    table.Append("}");
                }
                //table.Append("}");
            }
            table.Append("]");

            return table.ToString();
        }

    }
}
