using Ideal.Office.Excel;
using System.IO;
using System.Text;
namespace Ideal.Office.Web
{
    public partial class Import
    {
        public string ToJsonString(int titleIndex = 0)
        {
            var sht = Excel03.Sheet;
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

                    tempTd += sht.GetRow(rowIndex).GetCell(colIndex);
                    //列被合并之后此行将少输出colSpan-1个TD。
                    if (colSpan > 1)
                        colIndex += colSpan - 1;

                    if (colSpan > 1)
                        colIndex += colSpan - 1;
                    tempTd += ",";

                }
                var b = tempTd.Replace(",", "");
                if (!string.IsNullOrEmpty(b))
                {
                    table.Append("{");
                    table.Append(tempTd.Trim(','));
                    table.Append("}");
                }
            }
            table.Append("]");

            return table.ToString();
        }
    }
}
