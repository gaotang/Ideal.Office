namespace Ideal.Office.Web
{
    /// <summary>
    /// Excel数据列
    /// </summary>
    public class Column
    {
        public string ColumnName { get; set; }
        public string Property { get; set; }
        public bool IsPrimaryKey { get; set; }

        public bool Required { get; set; }
        public string DefValue { get; set; }
        public string DataType { get; set; }
        public string RefField { get; set; }
        public string RefTable { get; set; }
        public string RefConvertField { get; set; }
        public string RefFilter { get; set; }
        public decimal Min { get; set; }
        public decimal Max { get; set; }
        public int MaxLen { get; set; }
        public string Comment { get; set; }


        public bool IsEntry { get; set; }
    }
}
