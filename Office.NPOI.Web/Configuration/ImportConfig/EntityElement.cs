namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    public class EntityElement : ConfigurationElement
    {
        #region 字段
        private const string IdAttribute = "id";
        private const string ClassAttribute = "Classes";
        private const string ColumnAttribute = "Columnss";
        #endregion

        #region 属性
        /// <summary>
        /// 获取或设置元素值.
        /// </summary>
        [ConfigurationProperty(IdAttribute, IsRequired = true, IsKey = true)]
        public string Id
        {
            get { return (string)this[IdAttribute]; }
            set { this[IdAttribute] = value; }
        }

        [ConfigurationProperty(ClassAttribute, IsDefaultCollection = true)]
        public ClassElementCollection Class
        {
            get
            {
                return (ClassElementCollection)base[ClassAttribute];
            }
        }

        
        [ConfigurationProperty(ColumnAttribute, IsDefaultCollection = true)]
        public ColumnElementsCollection Columns
        {
            get
            {
                return (ColumnElementsCollection)base[ColumnAttribute];
            }
        }
        
        #endregion

        #region 方法

        #endregion
    }
}
