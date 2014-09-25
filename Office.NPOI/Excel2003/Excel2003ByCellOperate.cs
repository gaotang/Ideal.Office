namespace Ideal.Office.Excel
{
    using System.Collections.Generic;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /// <summary>
    /// 单元格操作
    /// </summary>
    public partial class Excel2003
    {


        #region 2.1.3 创建批注
        /// <summary>
        /// 创建批注
        /// </summary>
        public void CreateCellComment(ICell cell, string content)
        {
            if (!string.IsNullOrEmpty(content))
            {
                IDrawing patr = this.Sheet.CreateDrawingPatriarch();

                IComment comment1 = content.Length < 13
                    ? patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 4))
                    : patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 6));

                comment1.String = new HSSFRichTextString(content);
                comment1.Author = "系统管理员";
                cell.CellComment = comment1; 
            }
        }
        
        #endregion

        #region 2.2.2 合并单元格

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="cell">单元格</param>
        public void AddMergedRegion(ICell cell, int Rows = 1, int Cols = 1, double fontHeight = 15*15, HorizontalAlign postion = HorizontalAlign.Center)
        {
            if (Rows > 1 || Cols > 1)
            {
                var dic = new Dictionary<HorizontalAlign, HorizontalAlignment> { 
                    {HorizontalAlign.Center, HorizontalAlignment.Center },
                    {HorizontalAlign.Left, HorizontalAlignment.Left },
                    {HorizontalAlign.Right, HorizontalAlignment.Right },
                    {HorizontalAlign.Fill, HorizontalAlignment.Fill }
                };
                var rowNum = cell.Row.RowNum;
                var colNum = cell.ColumnIndex;

                var lastRow = rowNum + Rows - 1;
                var lastCol = colNum + Cols - 1;
                this.Sheet.AddMergedRegion(new CellRangeAddress(rowNum, lastRow, colNum, lastCol));

                
                //ICellStyle style = this.Workbook.CreateCellStyle();
                
                //style.Alignment = dic[postion];
                //style.VerticalAlignment = VerticalAlignment.Center;

                //IFont font = this.Workbook.CreateFont();
                //font.FontHeight = fontHeight;
                //style.SetFont(font);
                
                ICellStyle style = cell.CellStyle;
                //style.BorderBottom = BorderStyle.Thin;
                //style.BorderLeft = BorderStyle.Thin;
                //style.BorderRight = BorderStyle.Thin;
                //style.BorderTop = BorderStyle.Thin;
                //style.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Red.Index;

                style.Alignment = dic[postion];
                style.VerticalAlignment = VerticalAlignment.Center;

                cell.CellStyle = style;
            }
        }

        #endregion
    }
}
