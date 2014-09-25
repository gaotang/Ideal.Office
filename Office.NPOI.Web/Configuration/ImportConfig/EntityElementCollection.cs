namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    internal class EntityElementCollection : ConfigurationElementCollection
    {
        #region 字段
        private const string EntityName = "entity";
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

        public EntityElement this[int index]
        {
            get
            {
                return (EntityElement)BaseGet(index);
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
            return new EntityElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((EntityElement)element).Id;
        }

        #endregion
        
    }
}
