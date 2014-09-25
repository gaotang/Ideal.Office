namespace Ideal.Office.Excel
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using System.Linq;

    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.HPSF;

    /// <summary>
    /// 创建基本内容
    /// </summary>
    public partial class Excel2003
    {
        #region 3.1.1 创建Excel工作内容

        /// <summary>
        /// 创建Excel工作内容
        /// </summary>
        public HSSFWorkbook CreateWorkbook()
        {
            return new HSSFWorkbook();
        }

        public HSSFWorkbook CreateWorkbook(FileStream fs)
        {
            var poifs = new POIFSFileSystem(fs);
            return new HSSFWorkbook(poifs);
        }

        #endregion

        #region 3.1.2 创建DocumentSummaryInformation和SummaryInformation

        public void CreateDocumentSummaryInformation(string companyInfo)
        {
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = companyInfo;
            this.Workbook.DocumentSummaryInformation = dsi;
        }

        public void CreateSummaryInformation(string subjectInfo)
        {
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = subjectInfo;
            this.Workbook.SummaryInformation = si;
        }

        #endregion

        #region 3.1.3 设置页眉和页脚

        public void SetSheetHeader(string left = "", string center = "", string right = "")
        {
            if (!string.IsNullOrEmpty(left)) this.Sheet.Header.Left = left;
            if (!string.IsNullOrEmpty(center)) this.Sheet.Header.Center = center;
            if (!string.IsNullOrEmpty(right)) this.Sheet.Header.Right = right;
        }

        public void SetSheetFooter(string left = "", string center = "", string right = "")
        {
            if (!string.IsNullOrEmpty(left)) this.Sheet.Footer.Left = left;
            if (!string.IsNullOrEmpty(center)) this.Sheet.Footer.Center = center;
            if (!string.IsNullOrEmpty(right)) this.Sheet.Footer.Right = right;
        }

        #endregion

        #region 3.1.4 创建批注

        /// <summary>
        /// 创建批注
        /// </summary>
        /// <param name="dx1">      第1个单元格中x轴的偏移量</param>
        /// <param name="dy1">      第1个单元格中y轴的偏移量</param>
        /// <param name="dx2">      第2个单元格中x轴的偏移量</param>
        /// <param name="dy2">      第2个单元格中y轴的偏移量</param>
        /// <param name="col1">     第1个单元格的列号</param>
        /// <param name="row1">     第1个单元格的行号</param>
        /// <param name="col2">     第2个单元格的列号</param>
        /// <param name="row2">     第2个单元格的行号</param>
        /// <param name="content">  批注内容</param>
        /// <param name="author">   批注作者</param>
        /// <param name="cellRow">  单元格行</param>
        /// <param name="cellCol">  单元格列</param>
        public void CreateComment(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2, string content, string author, int cellRow, int cellCol)
        {
            var patr = this.Sheet.CreateDrawingPatriarch();
            var anchor = patr.CreateAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
            var comment1 = patr.CreateCellComment(anchor);

            comment1.String = new HSSFRichTextString(content);
            comment1.Author = author;

            var cell = this.Sheet.GetRow(cellRow).GetCell(cellCol);
            cell.CellComment = comment1;
        }

        #endregion

        #region 3.1.5 创建或获取 Excel工作单元

        public ISheet CreateSheet(string name)
        {
            return this.Workbook.CreateSheet(name);
        }

        public ISheet GetSheet(int index)
        {
            return this.Workbook.GetSheetAt(index);
        }

        public ISheet GetSheet(string name)
        {
            return this.Workbook.GetSheet(name);
        }

        public void SetSheetName(int index, string name)
        {
            this.Workbook.SetSheetName(index, name);
        }

        #endregion

        #region 3.1.6 创建或获取 Excel工作单元的行

        public IRow CreateRow(int index) {
            return this.Sheet.CreateRow(index);
        }

        public IRow GetRow(int index) {
            return this.Sheet.GetRow(index) ?? this.Sheet.CreateRow(index);
        }


        #endregion

        #region 3.1.7 创建或获取 Excel 单元格

        public ICell CreateCell(int rowIndex, int cellIndex) {
            return this.Sheet.GetRow(rowIndex).CreateCell(cellIndex);
        }


        public ICell GetCell(int rowIndex, int cellIndex)
        {
            return (this.Sheet.GetRow(rowIndex) 
                 ?? this.CreateRow(rowIndex)).GetCell(cellIndex)
                 ?? this.CreateCell(rowIndex, cellIndex); 
        }

        public void SetCellText(int rowIndex, int colIndex, int cellText)
        {
            var cell = this.GetCell(rowIndex, colIndex);
            
            cell.SetCellValue(cellText);
        }

        public void SetCellText(int rowIndex, int colIndex, string cellText)
        {
            var cell = this.GetCell(rowIndex, colIndex);
            cell.SetCellValue(cellText);
        }

        public void SetCellText(int rowIndex, int colIndex, double cellText)
        {
            var cell = this.GetCell(rowIndex, colIndex);
            cell.SetCellValue(cellText);
        }

        public void SetCellText(int rowIndex, int colIndex, DateTime cellText)
        {
            var cell = this.GetCell(rowIndex, colIndex);
            cell.SetCellValue(cellText);
        }

        /// <summary>
        /// 获取单元格的所有字符串形式
        /// </summary>
        /// <returns></returns>
        public string GetCellStrText(ICell cell) {
            var result = "";
            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        result = cell.DateCellValue.ToString();
                        break;
                    case CellType.Boolean:
                        result = cell.BooleanCellValue.ToString();
                        break;
                    case CellType.Error:
                        result = cell.ErrorCellValue.ToString();
                        break;
                    case CellType.Formula:
                        result = cell.CellFormula;
                        break;
                    case CellType.Numeric:
                        result = cell.NumericCellValue.ToString();
                        break;
                    case CellType.String:
                        result = cell.StringCellValue;
                        break;
                    case CellType.Unknown:
                        
                        break;
                    default:
                        break;
                }
            }
            return result;
        }

        #endregion

    }
}
