namespace Ideal.Office.Excel
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /// <summary>
    /// 使用Excel公式
    /// </summary>
    public partial class Excel2003
    {
        #region 2.3.2 用NPOI操作EXCEL－－SUM函数

        /// <summary>
        /// 用NPOI操作EXCEL
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="formula">相关的公式字符串</param>
        public void SetCellFormula(ICell cell, string formula)
        {
            cell.SetCellFormula(formula);
        }

        #endregion
    }
}
