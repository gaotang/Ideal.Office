namespace Ideal.Office.Web
{
    using Ideal.Office.Data;
    using Ideal.Office.Excel;
    using Ideal.Office.Web.Configuration;

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
        private static Excel2003 FillExcel(FileStream fs, ExcelEntity template, object[] dynamicCol, params object[] parameters)
        {
            return FillExcel(fs, template, dynamicCol, null, parameters);
        }
        /// <summary>
        /// 填充单元格
        /// </summary>
        /// <typeparam name="T">源数据类型</typeparam>
        /// <param name="poi">POI操作类</param>
        /// <param name="items">源数据</param>
        private static Excel2003 FillExcel(FileStream fs, ExcelEntity template, object[] dynamicCol, Dictionary<string, string> consts, params object[] parameters)
        {
            var poi = new Excel2003(fs);

            foreach (var item in template.Param)
            {
                poi.Sheet = poi.Workbook.GetSheetAt(item.Item1.SheetIndex);

                #region 1. 动态列的处理

                var query = item.Item3.Query;
                if (dynamicCol != null)
                    query = string.Format(query, dynamicCol);
                
                #endregion

                #region 2. 加载配置字段单元格

                if (item.Item2.Consts.Any())
                {
                    LoadConst(poi, item.Item2.Consts, consts);
                } 

                #endregion

                #region 3. 设置单元格

                //var dt = ExcelContainer.ToDataTable(query, parameters.Where(p => item.Item2.Param.Contains(p.ParameterName)).ToArray());
                var dt = new System.Data.DataTable();
                if (parameters.Length > 0)
                {
                    if (parameters[0] is Dictionary<string, object>)
                    {
                        var param = parameters[0] as Dictionary<string, object>;
                        var objParam = param.Where(p => item.Item3.Param.Contains(p.Key)).ToArray();
                        var keyParam = objParam.Select(p => p.Key).ToArray();
                        var valParam = objParam.Select(p => p.Value).ToArray();
                        for (int i = 0; i < keyParam.Length; i++)
                        {
                            query = query.Replace(keyParam[i], string.Concat("{", i, "}"));
                        }
                        dt = ExcelContainer.ToDataTable(query, valParam);
                    }
                    else dt = ExcelContainer.ToDataTable(query, parameters);
                }
                else {
                    dt = ExcelContainer.ToDataTable(query, parameters);
                }
                
                var rowCount = dt.Rows.Count;
                var columnCount = dt.Columns.Count;
                var dicCol = new Dictionary<int, int>();

                for (int i = 0; i < rowCount; i++)
                {

                    for (int j = 0; j < columnCount; j++)
                    {
                        var row = item.Item1.Row + i;
                        var col = item.Item1.Col + j;

                        #region 3.1 设置单元格内容
                        // 处理列的显示位置 如果有要处理的列和这是第一次处理
                        if (item.Item1.Childs.Any() && i == 0)
                        {
                            var cKey = string.Format("{0}", (j + 1));
                            if (item.Item1.Childs.ContainsKey(cKey))
	                        {
                                var cVal = item.Item1.Childs[cKey];
                                dicCol.Add(j, cVal.ToNumber() - 1);
	                        }
                        }
                        if (dicCol.ContainsKey(j))
                        {
                            col = dicCol[j];
                        }

                        var cellText = dt.Rows[i][j].ToString();
                        int intCell;
                        double dbCell;
                        DateTime dtCell;
                        if (int.TryParse(cellText, out intCell))
                             poi.SetCellText(row, col, intCell);
                        else if (double.TryParse(cellText, out dbCell))
                             poi.SetCellText(row, col, dbCell);
                        else if (DateTime.TryParse(cellText, out dtCell))
                             poi.SetCellText(row, col, dtCell);
                        else poi.SetCellText(row, col, cellText);

                        #endregion

                        #region 3.2 设置单元格相关计算公式

                        if (item.Item1.Formulas.Any())
                        {
                            foreach (var formula in item.Item1.Formulas)
                            {
                                if (formula.Postion.ToLower() == "top")
                                {
                                    if (i == 0 && (formula.ColNo == j + 1))
                                    {
                                        var cell = poi.GetCell(row - 1, col);
                                        poi.SetCellFormula(cell, formula.FormulaText);
                                    }
                                }
                                else if (formula.Postion.ToLower() == "bottom")
                                {
                                    if ((i == rowCount - 1) && formula.ColNo == j + 1)
                                    {
                                        var cell = poi.GetCell(row + 1, col);
                                        poi.SetCellFormula(cell, formula.FormulaText);
                                    }
                                }
                            }
                        }

                        #endregion
                    }
                } 

                #endregion

                #region 4. 合并单元格

                MergedCell(poi, item.Item1, dicCol);

                #endregion
            }

            return poi;
        }


        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="poi">Excel2003源数据</param>
        /// <param name="template">WebConfig配置数据源</param>
        /// <param name="dicCol">自定义列位置列表</param>
        private static void MergedCell(Excel2003 poi, SheetFlag sheetFlag, Dictionary<int, int> dicCol)
        {
            var sht = poi.Sheet;
            //取行Excel的最大行数
            int rowsCount = sht.PhysicalNumberOfRows;
            //为保证Table布局与Excel一样，这里应该取所有行中的最大列数（需要遍历整个Sheet）。
            //为少一交全Excel遍历，提高性能，我们可以人为把第0行的列数调整至所有行中的最大列数。

            //int colsCount = sht.GetRow(0).PhysicalNumberOfCells;
            int colsCount = poi.GetRow(0).PhysicalNumberOfCells;

            var mCols = sheetFlag.MergedCols;
            var isCommit = false;
            if (mCols.Any())
            {
                // 根据相关列处理数据
                var mColKey = sheetFlag.MergedColPrimaryKey;
                if (dicCol.Any() && dicCol.ContainsKey(mColKey - 1)) mColKey = dicCol[mColKey - 1] + 1;
                
                foreach (var item in mCols)
                {
                    var currentText = "";
                    // 主键列数据
                    var currentKeyText = "";
                    var currentIndex = 1;
                    ICell firstCell = sht.GetRow(0).GetCell(0); 

                    for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
                    {
                        if ((rowIndex + 1) >= sheetFlag.Row)
                        {
                            for (int colIndex = 0; colIndex < colsCount; colIndex++)
                            {
                                // 需要合并的列
                                if (item == colIndex + 1)
                                {
                                    // 获取当前列
                                    var tempItem = item - 1;
                                    if (dicCol.Any() && dicCol.ContainsKey(tempItem)) tempItem = dicCol[tempItem];
                                    var cell = poi.GetRow(rowIndex).GetCell(tempItem);

                                    if (cell != null)
                                    {
                                        if (cell.CellType == CellType.String)
                                        {
                                            #region 1.1 单元格内容为 字符串
                                            if (!string.IsNullOrEmpty(cell.StringCellValue))
                                            {
                                                if (string.IsNullOrEmpty(currentText))
                                                {
                                                    currentText = cell.StringCellValue.Trim();
                                                    if (mColKey > 0)
                                                    {
                                                        var currentKey = poi.GetRow(rowIndex).GetCell(mColKey - 1);
                                                        currentKeyText = currentKey.CellType == CellType.Numeric ? currentKey.NumericCellValue.ToString().Trim() : currentKey.StringCellValue;
                                                    }
                                                    firstCell = cell;
                                                    currentIndex = 1;
                                                }
                                                else if (currentText != cell.StringCellValue.Trim())
                                                {
                                                    poi.AddMergedRegion(firstCell, currentIndex);
                                                    isCommit = false;
                                                    firstCell = cell;
                                                    currentText = cell.StringCellValue.Trim();
                                                    if (mColKey > 0)
                                                    {
                                                        var currentKey = poi.GetRow(rowIndex).GetCell(mColKey - 1);
                                                        currentKeyText = currentKey.CellType == CellType.Numeric ? currentKey.NumericCellValue.ToString().Trim() : currentKey.StringCellValue;
                                                    }
                                                    currentIndex = 1;
                                                }
                                                else
                                                {
                                                    if (mColKey > 0)
                                                    {
                                                        var currentKey = poi.GetRow(rowIndex).GetCell(mColKey - 1);
                                                        var tempKeyText = currentKey.CellType == CellType.Numeric ? currentKey.NumericCellValue.ToString().Trim() : currentKey.StringCellValue;
                                                        
                                                        if (currentKeyText != tempKeyText)
                                                        {
                                                            poi.AddMergedRegion(firstCell, currentIndex);
                                                            isCommit = false;
                                                            firstCell = cell;
                                                            currentKeyText = tempKeyText;
                                                            currentIndex = 0;
                                                        }
                                                    }

                                                    currentIndex++;
                                                    isCommit = true;
                                                }
                                            }
                                            #endregion
                                        }
                                        else if (cell.CellType == CellType.Numeric)
                                        {
                                            #region 1.2 单元格内容为 数字型
                                            if (cell.NumericCellValue != double.NaN)
                                            {
                                                if (string.IsNullOrEmpty(currentText))
                                                {
                                                    currentText = cell.NumericCellValue.ToString().Trim();
                                                    if (mColKey > 0)
                                                    {
                                                        var currentKey = poi.GetRow(rowIndex).GetCell(mColKey - 1);
                                                        currentKeyText = currentKey.CellType == CellType.Numeric ? currentKey.NumericCellValue.ToString().Trim() : currentKey.StringCellValue;
                                                    }
                                                    firstCell = cell;
                                                    currentIndex = 1;
                                                }
                                                else if (currentText != cell.NumericCellValue.ToString().Trim())
                                                {
                                                    poi.AddMergedRegion(firstCell, currentIndex);
                                                    isCommit = false;
                                                    firstCell = cell;
                                                    currentText = cell.NumericCellValue.ToString().Trim();
                                                    if (mColKey > 0)
                                                    {
                                                        var currentKey = poi.GetRow(rowIndex).GetCell(mColKey - 1);
                                                        currentKeyText = currentKey.CellType == CellType.Numeric ? currentKey.NumericCellValue.ToString().Trim() : currentKey.StringCellValue;
                                                    }
                                                    currentIndex = 1;
                                                }
                                                else
                                                {
                                                    if (mColKey > 0)
                                                    {
                                                        var currentKey = poi.GetRow(rowIndex).GetCell(mColKey - 1);
                                                        var tempKeyText = currentKey.CellType == CellType.Numeric ? currentKey.NumericCellValue.ToString().Trim() : currentKey.StringCellValue;
                    
                                                        if (currentKeyText != tempKeyText)
                                                        {
                                                            poi.AddMergedRegion(firstCell, currentIndex);
                                                            isCommit = false;
                                                            firstCell = cell;
                                                            currentKeyText = tempKeyText;
                                                            currentIndex = 0;
                                                        }
                                                    }

                                                    currentIndex++;
                                                    isCommit = true;
                                                }
                                            }
                                            #endregion
                                        }
                                        
                                    }
                                }
                            }
                            if (isCommit)
                            {
                                isCommit = false;
                                poi.AddMergedRegion(firstCell, currentIndex);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 加载常量处理数据
        /// </summary>
        /// <param name="poi">Excel2003操作类</param>
        /// <param name="childs">配置文件常量字典</param>
        /// <param name="consts">用户操作常量字典</param>
        private static void LoadConst(Excel2003 poi, List<System.Configuration.NameValueConfigurationElement> childs, Dictionary<string, string> consts)
        {
            var sht = poi.Sheet;
            //取行Excel的最大行数
            int rowsCount = sht.PhysicalNumberOfRows;
            //为保证Table布局与Excel一样，这里应该取所有行中的最大列数（需要遍历整个Sheet）。
            //为少一交全Excel遍历，提高性能，我们可以人为把第0行的列数调整至所有行中的最大列数。
            int colsCount = sht.GetRow(0).PhysicalNumberOfCells;

            int colSpan;
            int rowSpan;
            bool isByRowMerged;

            for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
            {
                for (int colIndex = 0; colIndex < colsCount; colIndex++)
                {
                    GetTdMergedInfo(poi, rowIndex, colIndex, out colSpan, out rowSpan, out isByRowMerged);
                    //如果已经被行合并包含进去了就不输出TD了。
                    //注意被合并的行或列不输出的处理方式不一样，见下面一处的注释说明了列合并后不输出TD的处理方式。
                    if (isByRowMerged)
                    {
                        continue;
                    }

                    //var cell = sht.GetRow(rowIndex).GetCell(colIndex);
                    var cell = poi.GetRow(rowIndex).GetCell(colIndex);
                    if (cell != null)
                    {
                        var cellText = poi.GetCellStrText(cell);
                        foreach (var item in childs)
                        {
                            if (consts != null)
                            {
                                foreach (var const1 in consts)
                                {
                                    if (item.Value == const1.Key)
                                    {
                                        if (cellText.Contains(item.Name))
                                        {
                                            cellText = poi.GetCellStrText(cell);
                                            var newValue = cellText.Replace(item.Name, const1.Value);
                                            cell.SetCellValue(newValue);
                                        }
                                    }
                                }
                            }

                            if (cell.CellType == CellType.String)
                            {
                                if (cell.StringCellValue.Contains(item.Name))
                                {
                                    var newValue = cellText.Replace(item.Name, item.Value);
                                    cell.SetCellValue(newValue);
                                }
                            }
                            else if (cell.CellType == CellType.Numeric) {
                                if (cell.NumericCellValue.ToString().Contains(item.Name))
                                {
                                    var newValue = cellText.Replace(item.Name, item.Value);
                                    cell.SetCellValue(newValue);
                                }
                            }
                            
                        } 
                    }

                    //列被合并之后此行将少输出colSpan-1个TD。
                    if (colSpan > 1)
                        colIndex += colSpan - 1;

                }
            }
        }

        //public static MemoryStream Down(string templateName, params object[] parameters)
        //{
        //    return Down(templateName, null, null, parameters);
        //}

        public static MemoryStream Down(string templateName, Dictionary<string, string> consts = null)
        {
            return Down(templateName, consts, null, new object[0]);
        }

        public static MemoryStream Down(string templateName, Dictionary<string, string> consts = null, params object[] parameters)
        {
            return Down(templateName, consts, null, parameters);
        }
        public static MemoryStream Down(string templateName, Dictionary<string, string> consts = null, object[] dynamicCol = null, params object[] parameters)
        {
            ExcelEntity entity;
            if (Import.GetXmlConfig().TryGetValue(templateName, out entity))
            {
                //var name = excelName + DateTime.Now.ToString("(yyyy年MM月dd日hh时mm分ss秒)") + (entity.Type == ExcelType.Excel2003 ? ".xls" : ".xlsx");
                using (FileStream fs = new FileStream(entity.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    Excel = FillExcel(fs, entity, dynamicCol, consts, parameters);

                    using (MemoryStream Stream = new MemoryStream())
                    {
                        Excel.Workbook.Write(Stream);
                        return Stream;
                        //HttpContext.Current.Response.ExportExcel(Stream, name);
                    }
                }
            }
            else
            {
                return null;
            }

        }


    }
}
