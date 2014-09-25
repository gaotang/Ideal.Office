namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    internal class TemplateElementCollection : ConfigurationElementCollection
    {
        #region 字段
        private const string NameAttribute = "key";
        private const string PathAttribute = "value";
        private const string TypeAttribute = "type";
        private const string SqlAttribute = "sql";

        private const string ConstsAttribute = "consts";
        #endregion

        #region 属性
        [ConfigurationProperty(NameAttribute, IsKey = true, IsRequired = true)]
        public string Name
        {
            get
            {
                return (string)base[NameAttribute];
            }
            set
            {
                base[NameAttribute] = value;
            }
        }

        [ConfigurationProperty(PathAttribute, IsRequired = true)]
        public string Path
        {
            get
            {
                return (string)base[PathAttribute];
            }
            set
            {
                base[PathAttribute] = value;
            }
        }

        [ConfigurationProperty(TypeAttribute, IsRequired = true)]
        public string Type
        {
            get
            {
                return (string)base[TypeAttribute];
            }
            set
            {
                base[TypeAttribute] = value;
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new SqlItemElementCollection();
        }
        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((SqlItemElementCollection)element).Key;
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
                return SqlAttribute;
            }
        }

        public SqlItemElementCollection this[int index]
        {
            get
            {
                return (SqlItemElementCollection)BaseGet(index);
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

        [ConfigurationProperty(ConstsAttribute, IsDefaultCollection = false)]
        public System.Configuration.NameValueConfigurationCollection Consts
        {
            get
            {
                return (NameValueConfigurationCollection)base[ConstsAttribute];
            }
            set
            {
                base[ConstsAttribute] = value;
            }
        }

        #endregion
    }
}
