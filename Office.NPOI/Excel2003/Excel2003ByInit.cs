namespace Ideal.Office.Excel
{
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /// <summary>
    /// 初始化Excel2003
    /// </summary>
    public partial class Excel2003
    {
        #region 1. 属性

        /// <summary>
        /// Excel工作内容
        /// </summary>
        public HSSFWorkbook Workbook { get; set; }

        /// <summary>
        /// Excel工作单元
        /// </summary>
        public ISheet Sheet { get; set; }

        /// <summary>
        /// 当前行
        /// </summary>
        public int CurrentRow { get; set; }

        /// <summary>
        /// 当前列
        /// </summary>
        public int CurrentCol { get; set; }

        /// <summary>
        /// 总列数
        /// </summary>
        public int ColsCount { get; set; }

        /// <summary>
        /// 计算公式操作类
        /// </summary>
        public HSSFFormulaEvaluator Formula { get; private set; } 

        #endregion

        #region 2. 构造函数

        public Excel2003(string sheetName) {
            this.Workbook = this.CreateWorkbook();
            this.Sheet = this.GetSheet(sheetName) ?? this.CreateSheet(sheetName);
        }

        public Excel2003(FileStream fs)
        {
            this.Workbook = this.CreateWorkbook(fs);
        }

        public Excel2003(HSSFWorkbook workbook, ISheet sheet)
        {
            this.Workbook = workbook;
            this.Sheet = sheet;
        }

        public Excel2003(HSSFWorkbook workbook, ISheet sheet, int colsCount)
            : this(workbook, sheet)
        {
            this.ColsCount = colsCount;
        }

        #endregion

        
    }
}
