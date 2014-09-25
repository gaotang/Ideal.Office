namespace Ideal.Office.Web
{
    using System.Collections.Generic;

    public class SheetFlag
    {
        /// <summary>
        /// Excel生成行起点
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// Excel生成列起点
        /// </summary>
        public int Col { get; set; }
        /// <summary>
        /// Excel表格索引页
        /// </summary>
        public int SheetIndex { get; set; }

        /// <summary>
        /// Excel表格相关计算公式
        /// </summary>
        public List<FormulaEntity> Formulas { get; set; }

        
        /// <summary>
        /// Excel表格中需要合并的列的依据列
        /// </summary>
        public int MergedColPrimaryKey { get; set; }

        /// <summary>
        /// Excel表格中需要合并的列
        /// </summary>
        public int[] MergedCols { get; set; }

        /// <summary>
        /// Excel表格中需要合并的行的依据行
        /// </summary>
        public int MergedRowPrimaryKey { get; set; }

        /// <summary>
        /// Excel表格中需要合并的行
        /// </summary>
        public int[] MergedRows { get; set; }

        /// <summary>
        /// 相关的Excel常量处理
        /// </summary>
        public Dictionary<string, string> Childs { get; set; }
    }
}
