namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;
    internal class CatalogElementCollection : ConfigurationElementCollection
    {
        #region 字段
        private const string CatalogAttribute = "catalog";
        #endregion

        #region 属性
        protected override ConfigurationElement CreateNewElement()
        {
            return new CatalogElement();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((CatalogElement)element).Key;
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
                return CatalogAttribute;
            }
        }

        public CatalogElement this[int index]
        {
            get
            {
                return (CatalogElement)BaseGet(index);
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
    }
}
