namespace Ideal.Office.Web
{
    using System;
    using System.Collections.Generic;

    using Ideal.Office.Excel;

    public class ExcelEntity
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public ExcelType Type { get; set; }
        public List<Tuple<SheetFlag, ConstFlag, SqlFlag>> Param { get; set; }

    }
}
