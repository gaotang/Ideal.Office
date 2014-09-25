namespace Ideal.Office.Web
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    using System;
    using System.Collections.Generic;
    using System.Reflection;

    /// <summary>
    /// Excel数据导入
    /// </summary>
    public class ExcelDataImport
    {
        public Regulation regulation { get; set; }

        public string excelFileName { get; set; }

        HSSFSheet sheet { get; set; }

        public ExcelDataImport(HSSFSheet _sheet)
        {
            this.sheet = _sheet;
        }

        private void Execute()
        {
            int rowsCount = sheet.PhysicalNumberOfRows;
            Object headEntity = null, entryEntity = null;
            List<object> entries = null;
            for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
            {
                if (headEntity == null)
                {
                    //构建表头实体
                    headEntity = GetHeadEntity(rowIndex);
                    //构建表体实体
                    entries = new List<object>();
                    entryEntity = GetHeadEntity(rowIndex);
                    entries.Add(entryEntity);
                }
                //检查表头关键字段值是否与下一行相等，如果相等则当同一张单据处理，只用构建表体；如果不等说明是一张新的单据，要先把当前单据保存。
                if (IsSameToNextRow(rowIndex))
                {
                    entryEntity = GetHeadEntity(rowIndex);
                    entries.Add(entryEntity);
                }
                else
                {
                    //设置分录属性的值
                    Type head = Type.GetType(regulation.HeadClass);
                    PropertyInfo entryProperty = head.GetProperty(regulation.EntryProperty);
                    //entryProperty.SetValue(headEntity, entries);
                    entryProperty.SetValue(headEntity, entries, null);

                    Type clsMethod = Type.GetType(regulation.MethodClass);
                    Object service = Activator.CreateInstance(clsMethod);
                    MethodInfo saveMethod = clsMethod.GetMethod(regulation.SaveMethod);
                    MethodInfo exValidateMethod = clsMethod.GetMethod(regulation.ExValidateMethod);
                    if (exValidateMethod != null)
                        //执行附加的检查方法
                        exValidateMethod.Invoke(service, new object[] { headEntity });
                    if (saveMethod != null)
                        //执行保存方法
                        exValidateMethod.Invoke(service, new object[] { headEntity });
                    //重置为一张新的单据
                    headEntity = null;
                    entries = null;
                }
            }
        }

        private bool IsSameToNextRow(int rowIndex)
        {
            throw new NotImplementedException();
        }

        private Object GetHeadEntity(int rowIndex)
        {
            IRow row = sheet.GetRow(rowIndex); ;
            ICell cell;
            Type cls = Type.GetType(regulation.HeadClass);
            Object obj = Activator.CreateInstance(cls);
            PropertyInfo pi;
            Object val;
            int colIndex;
            foreach (Column col in regulation.Columns)
            {
                if (!col.IsEntry)
                {
                    //根据列名找出Excel中对应的列序号
                    colIndex = GetExcelColumnIndex(col.ColumnName);
                    cell = row.GetCell(colIndex);
                    //根据Column的数据类型取出Cell中的内容。如果为空将取Column对象中定义的默认值
                    val = GetCellValue(cell, col.DataType);
                    //获取实体类中的属性
                    pi = cls.GetProperty(col.Property);
                    //给实体类中的属性赋值
                    //pi.SetValue(obj, val);
                    pi.SetValue(obj, val, null);
                }
            }
            return obj;
        }

        /// <summary>
        /// 根据Column的数据类型取出Cell中的内容。如果为空将取Column对象中定义的默认值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="dataType">数据类型</param>
        private object GetCellValue(ICell cell, string dataType)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 根据列名找出Excel中对应的列序号
        /// </summary>
        /// <param name="columnName">列名</param>
        private int GetExcelColumnIndex(string columnName)
        {
            return 0;
        }

        public object GetEntryEntity(int rowIndex)
        {
            return null;
        }
    }
}
