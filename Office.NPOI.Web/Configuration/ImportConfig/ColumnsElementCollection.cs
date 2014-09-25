namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    public class ColumnElementsCollection : ConfigurationElementCollection
    {

        #region 字段
        private const string EntityName = "Columns";

        #endregion

        #region 属性
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

        public ColumnElementCollection this[int index]
        {
            get
            {
                return (ColumnElementCollection)BaseGet(index);
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
            return new ColumnElementCollection();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ColumnElementCollection)element).Key;
        }

        #endregion
    }
}
