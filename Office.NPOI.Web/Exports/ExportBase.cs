namespace Ideal.Office.Web
{
    using Ideal.Office.Data;
    using Ideal.Office.Excel;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using System.IO;
    using System.Linq;
    using System.Web;
    public partial class Export
    {
        public static Excel2003 Excel { get; private set; }


        /// <summary>
        ///  获取Table某个TD合并的列数和行数等信息。与Excel中对应Cell的合并行数和列数一致。
        /// </summary>
        /// <param name="rowIndex">行号</param>
        /// <param name="colIndex">列号</param>
        /// <param name="colspan">TD中需要合并的行数</param>
        /// <param name="rowspan">TD中需要合并的列数</param>
        /// <param name="rowspan">此单元格是否被某个行合并包含在内。如果被包含在内，将不输出TD。</param>
        /// <returns></returns>
        private static void GetTdMergedInfo(int rowIndex, int colIndex, out int colspan, out int rowspan, out bool isByRowMerged)
        {
            var sht = Excel.Sheet;
            colspan = 1;
            rowspan = 1;
            isByRowMerged = false;
            int regionsCuont = sht.NumMergedRegions;
            //Region region;
            CellRangeAddress region;

            for (int i = 0; i < regionsCuont; i++)
            {
                //region = sht.GetMergedRegionAt(i);
                //if (region.RowFrom == rowIndex && region.ColumnFrom == colIndex)
                //{
                //    colspan = region.ColumnTo - region.ColumnFrom + 1;
                //    rowspan = region.RowTo - region.RowFrom + 1;
                //    return;
                //}
                //else if (rowIndex > region.RowFrom && rowIndex <= region.RowTo && colIndex >= region.ColumnFrom && colIndex <= region.ColumnTo)
                //{
                //    isByRowMerged = true;
                //}
                region = sht.GetMergedRegion(i);
                if (region.FirstRow == rowIndex && region.FirstColumn == colIndex)
                {
                    colspan = region.LastColumn - region.FirstColumn + 1;
                    rowspan = region.LastRow - region.FirstRow + 1;
                    return;
                }
                else if (rowIndex > region.FirstRow && rowIndex <= region.LastRow && colIndex >= region.FirstColumn && colIndex <= region.LastColumn)
                {
                    isByRowMerged = true;
                }
            }
        }


        private static void GetTdMergedInfo(Excel2003 poi, int rowIndex, int colIndex, out int colspan, out int rowspan, out bool isByRowMerged)
        {
            var sht = poi.Sheet;
            colspan = 1;
            rowspan = 1;
            isByRowMerged = false;
            int regionsCuont = sht.NumMergedRegions;
            //Region region;
            CellRangeAddress region;

            for (int i = 0; i < regionsCuont; i++)
            {
                region = sht.GetMergedRegion(i);
                if (region.FirstRow == rowIndex && region.FirstColumn == colIndex)
                {
                    colspan = region.LastColumn - region.FirstColumn + 1;
                    rowspan = region.LastRow - region.FirstRow + 1;
                    return;
                }
                else if (rowIndex > region.FirstRow && rowIndex <= region.LastRow && colIndex >= region.FirstColumn && colIndex <= region.LastColumn)
                {
                    isByRowMerged = true;
                }
            }
        }

    }
}
