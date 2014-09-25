namespace Ideal.Office.Excel
{
    using System;
    using System.Collections.Generic;

    public static class EnumerableExtensions
    {
        #region 方法

        public static void ForEach<T>(this IEnumerable<T> items, Action<T> action)
        {
            foreach (T item in items) action(item);
        }

        ///// <summary>
        ///// 生成相关Excel
        ///// </summary>
        //public static void ToExcel<T>(this IEnumerable<T> items, string templateName, string excelName = "", ExcelType type = ExcelType.Excel2003)
        //{
        //    ExcelApp.Export<T>(items, excelName, templateName, type);
        //}

        #endregion
    }
}
