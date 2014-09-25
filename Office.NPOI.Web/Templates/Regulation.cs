namespace Ideal.Office.Web
{
    /// <summary>
    /// Template导入导出 方法管理类
    /// </summary>
    public class Regulation
    {
        public string HeadClass { get; set; }

        public string EntryProperty { get; set; }

        public string EntryClass { get; set; }

        public string SaveMethod { get; set; }

        public string ExValidateMethod { get; set; }

        public string MethodClass { get; set; }

        public Column[] Columns { get; set; }
    }
}
