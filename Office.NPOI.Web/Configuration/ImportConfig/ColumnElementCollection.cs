namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    public class ColumnElementCollection : ConfigurationElementCollection
    {

        #region 字段
        private const string EntityName = "Column";
        private const string KeyAttribute = "key";
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

        public ColumnElement this[int index]
        {
            get
            {
                return (ColumnElement)BaseGet(index);
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
            return new ColumnElement();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ColumnElement)element).Property;
        }

        #endregion
    }
}
