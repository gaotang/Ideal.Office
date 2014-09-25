namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    internal class SqlItemElementCollection : ConfigurationElementCollection
    {
        #region 字段
        private const string KeyAttribute = "key";
        private const string ValueAttribute = "value";
        private const string StartAttribute = "start";
        private const string AutoColAttribute = "autoCol";
        private const string ParamAttribute = "param";
        private const string SheetIndexAttribute = "sheetIndex";
        private const string FormulaAttribute = "formula";
        private const string MergedColPrimaryKeyAttribute = "mergedColPrimaryKey";
        private const string MergedColsAttribute = "mergedCols";
        private const string MergedRowPrimaryKeyAttribute = "mergedRowPrimaryKey";
        private const string MergedRowsAttribute = "mergedRows";
        

        private const string EntityName = "add";
        
        #endregion

        #region 属性
        /// <summary>
        /// 获取或设置目录名称.
        /// </summary>
        [ConfigurationProperty(KeyAttribute, IsRequired = true, IsKey = true)]
        public string Key
        {
            get { return (string)this[KeyAttribute]; }
            set { this[KeyAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置目录路径.
        /// </summary>
        [ConfigurationProperty(ValueAttribute, IsRequired = true)]
        public string Value
        {
            get { return (string)this[ValueAttribute]; }
            set { this[ValueAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置Excel模板的开始点.
        /// </summary>
        [ConfigurationProperty(StartAttribute, IsRequired = true)]
        public string Start
        {
            get { return (string)this[StartAttribute]; }
            set { this[StartAttribute] = value; }
        }

        [ConfigurationProperty(AutoColAttribute, IsRequired = false)]
        public string AutoCol
        {
            get { return (string)this[AutoColAttribute]; }
            set { this[AutoColAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置SQL语句的参数.
        /// </summary>
        [ConfigurationProperty(ParamAttribute, IsRequired = false)]
        public string Param
        {
            get { return (string)this[ParamAttribute]; }
            set { this[ParamAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置需要操作的Excel页面索引(从0开始).
        /// </summary>
        [ConfigurationProperty(SheetIndexAttribute, IsRequired = true)]
        public int SheetIndex
        {
            get { return (int)this[SheetIndexAttribute]; }
            set { this[SheetIndexAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置当前页面单元格的计算公式.
        /// </summary>
        [ConfigurationProperty(FormulaAttribute, IsRequired = false)]
        public string Formula
        {
            get { return (string)this[FormulaAttribute]; }
            set { this[FormulaAttribute] = value; }
        }
        
        /// <summary>
        /// 获取或设置Excel中需要合并的列主键.
        /// </summary>
        [ConfigurationProperty(MergedColPrimaryKeyAttribute, IsRequired = false)]
        public int MergedColPrimaryKey
        {
            get { return (int)this[MergedColPrimaryKeyAttribute]; }
            set { this[MergedColPrimaryKeyAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置Excel中需要合并的列集合.
        /// </summary>
        [ConfigurationProperty(MergedColsAttribute, IsRequired = false)]
        public string MergedCols
        {
            get { return (string)this[MergedColsAttribute]; }
            set { this[MergedColsAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置Excel中需要合并的列主键.
        /// </summary>
        [ConfigurationProperty(MergedRowPrimaryKeyAttribute, IsRequired = false)]
        public int MergedRowPrimaryKey
        {
            get { return (int)this[MergedRowPrimaryKeyAttribute]; }
            set { this[MergedRowPrimaryKeyAttribute] = value; }
        }


        /// <summary>
        /// 获取或设置Excel中需要合并的行集合.
        /// </summary>
        [ConfigurationProperty(MergedRowsAttribute, IsRequired = false)]
        public string MergedRows
        {
            get { return (string)this[MergedRowsAttribute]; }
            set { this[MergedRowsAttribute] = value; }
        }

        

        public override ConfigurationElementCollectionType CollectionType
        {
            get
            {
                return ConfigurationElementCollectionType.BasicMap;
            }
        }
        protected override string ElementName
        {
            get
            {
                return EntityName;
            }
        }

        public AddItemElement this[int index]
        {
            get
            {
                return (AddItemElement)BaseGet(index);
            }
            set
            {
                if (BaseGet(index) != null)
                {
                    BaseRemoveAt(index);
                }
                BaseAdd(index, value);
            }
        }
        #endregion

        #region 方法
        protected override ConfigurationElement CreateNewElement()
        {
            return new AddItemElement();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((AddItemElement)element).Key;
        }

        #endregion
    }
}
