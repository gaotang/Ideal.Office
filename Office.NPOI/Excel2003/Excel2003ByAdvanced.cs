namespace Ideal.Office.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /// <summary>
    /// 高级功能
    /// </summary>
    public partial class Excel2003
    {
        #region 3.6.1 调整表单显示比例

        /// <summary>
        /// 调整表单显示比例
        /// </summary>
        /// <param name="numerator">放大比例的分子</param>
        /// <param name="denominator">放大比例的分母</param>
        public void SetZoom(int numerator, int denominator)
        {
            this.Sheet.SetZoom(numerator, denominator);   
        }

        #endregion

        #region 3.6.8 生成下拉列表

        /// <summary>
        /// Excel2003
        /// </summary>
        /// <param name="contents">下拉列表字符串数组，不能超过255个字符</param>
        /// <param name="firstCol">第一列</param>
        /// <param name="lastCol">最后一列</param>
        /// <param name="fistRow">第一行</param>
        /// <param name="lastRow">最后一行</param>
        public void CreateDropDownList(string[] contents,int firstCol,int lastCol, int fistRow, int lastRow = 65535)
        {
            CellRangeAddressList regions = new CellRangeAddressList(fistRow, lastRow, firstCol, lastCol);
            DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(contents);
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            this.Sheet.AddValidationData(dataValidate);
        }

        #endregion

        #region 3.6.9 隐藏列

        public void SetColumnHidden(string colNum)
        {
            this.Sheet.SetColumnHidden(colNum.ToNumber(), true);
        }

        public void SetColumnHidden(int colNum)
        {
            this.Sheet.SetColumnHidden(colNum, true);
        }

        public void SetColumnHidden(int firstCol, int lastCol, int[] Array){
            var xIndex = 1;
            var xDic = new Dictionary<int, int>();
            for (int x = firstCol; x <= lastCol; x++)
            {
                xDic.Add(xIndex, x);
                xIndex++;
            }
            var AllMonth = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            var hideMonth = AllMonth.Except(Array).ToList();
            xIndex = 0;
            for (int y = 0; y < hideMonth.Count; y++)
            {
                SetColumnHidden(xDic[hideMonth[y]]);
            }
        }
        #endregion

    }
}
