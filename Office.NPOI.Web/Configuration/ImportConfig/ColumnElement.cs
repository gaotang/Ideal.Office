namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    public class ColumnElement : ConfigurationElement
    {
        #region 字段
        private const string PropertyAttribute = "Property";
        private const string IsEntryAttribute = "IsEntry";
        private const string IsPrimaryKeyAttribute = "IsPrimaryKey";
        private const string ColumnNameAttribute = "ColumnName";
        private const string RequriedAttribute = "Required";
        private const string DataTypeAttribute = "DataType";
        private const string RefConfigAttribute = "RefConfig";
        private const string DefValueAttribute = "DefValue";
        private const string MinAttribute = "Min";
        private const string MaxAttribute = "Max";
        private const string MaxLenAttribute = "MaxLen";
        private const string CommentAttribute = "Comment";
        private const string RowsAttribute = "Rows";
        private const string ColsAttribute = "Cols";
        #endregion

        #region 属性
        /// <summary>
        /// 获取或设置实体类中对应的属性名.
        /// </summary>
        [ConfigurationProperty(PropertyAttribute, IsRequired = true, IsKey = true)]
        public string Property
        {
            get { return (string)this[PropertyAttribute]; }
            set { this[PropertyAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置是否表体字段。默认为0，不是表体字段.
        /// </summary>
        [ConfigurationProperty(IsEntryAttribute, IsRequired = false)]
        public int IsEntry
        {
            get { return (int)this[IsEntryAttribute]; }
            set { this[IsEntryAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置是否唯一字段。因一个订单头可能对应多个订单体，即Excel中的多条记录对应一张订单。导入时将根据IsPrimaryKey=1的字段值来确定哪些行属于一张订单.
        /// </summary>
        [ConfigurationProperty(IsPrimaryKeyAttribute, IsRequired = false)]
        public int IsPrimaryKey
        {
            get { return (int)this[IsPrimaryKeyAttribute]; }
            set { this[IsPrimaryKeyAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置是否唯一字段。因一个订单头可能对应多个订单体，即Excel中的多条记录对应一张订单。导入时将根据IsPrimaryKey=1的字段值来确定哪些行属于一张订单.
        /// </summary>
        [ConfigurationProperty(ColumnNameAttribute, IsRequired = false)]
        public string ColumnName
        {
            get { return (string)this[ColumnNameAttribute]; }
            set { this[ColumnNameAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置是否必录项。默认0，不必录.
        /// </summary>
        [ConfigurationProperty(RequriedAttribute, IsRequired = false)]
        public int Requried
        {
            get { return (int)this[RequriedAttribute]; }
            set { this[RequriedAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置生成引入模板时将根据此生成对应列的有效性验证。如果是引用类型（DataType=ref），则会根据RefConfig生成下拉列表及保存时进行相应的转换（如将客户名称转换为客户ID，因为数据库中存的为ID，Excel中显示的为名称）.
        /// </summary>
        [ConfigurationProperty(DataTypeAttribute, IsRequired = false)]
        public string DataType
        {
            get { return (string)this[DataTypeAttribute]; }
            set { this[DataTypeAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置引用类型的配置信息。引用类型一般需要转换后保存，故配置项依次为 转换后的字段;转换时查找的表;转换时用到的比较字段;过滤条件.
        /// </summary>
        [ConfigurationProperty(RefConfigAttribute, IsRequired = false)]
        public string RefConfig
        {
            get { return (string)this[RefConfigAttribute]; }
            set { this[RefConfigAttribute] = value; }
        }

        /// <summary>
        /// 获取或设置默认值。支持多种默认值：$F{XXX}表示取系统级的变量，当然系统中要首先存在这些变量；$E{a.b}表示取其它对象的相关属性；$C{XXX}表示默认值为常量；$S{XXX}表示通过SQL取默认值等
        /// </summary>
        [ConfigurationProperty(DefValueAttribute, IsRequired = false)]
        public string DefValue
        {
            get { return (string)this[DefValueAttribute]; }
            set { this[DefValueAttribute] = value; }
        }

        /// <summary>
        /// 最少值
        /// </summary>
        [ConfigurationProperty(MinAttribute, IsRequired = false)]
        public int Min
        {
            get { return (int)this[MinAttribute]; }
            set { this[MinAttribute] = value; }
        }

        /// <summary>
        /// 最大值
        /// </summary>
        [ConfigurationProperty(MaxAttribute, IsRequired = false)]
        public int Max
        {
            get { return (int)this[MaxAttribute]; }
            set { this[MaxAttribute] = value; }
        }

        /// <summary>
        /// 最大长度
        /// </summary>
        [ConfigurationProperty(MaxLenAttribute, IsRequired = false)]
        public int MaxLen
        {
            get { return (int)this[MaxLenAttribute]; }
            set { this[MaxLenAttribute] = value; }
        }

        /// <summary>
        /// 生成引入模板时列头的批注
        /// </summary>
        [ConfigurationProperty(CommentAttribute, IsRequired = false)]
        public string Comment
        {
            get { return (string)this[CommentAttribute]; }
            set { this[CommentAttribute] = value; }
        }

        /// <summary>
        /// 跨列数量
        /// </summary>
        [ConfigurationProperty(ColsAttribute, IsRequired = false)]
        public int Cols
        {
            get { return (int)this[ColsAttribute]; }
            set { this[ColsAttribute] = value; }
        }

        /// <summary>
        /// 跨行数量
        /// </summary>
        [ConfigurationProperty(RowsAttribute, IsRequired = false)]
        public int Rows
        {
            get { return (int)this[RowsAttribute]; }
            set { this[RowsAttribute] = value; }
        }
        #endregion


        #region 方法
        
        #endregion

    }
}
