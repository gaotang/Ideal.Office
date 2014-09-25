namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    public class ClassElementCollection : ConfigurationElementCollection
    {

        #region 字段
        private const string EntityName = "Class";
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

        public ClassElement this[int index]
        {
            get
            {
                return (ClassElement)BaseGet(index);
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
            return new ClassElement();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ClassElement)element).Key;
        }

        #endregion
    }

}
