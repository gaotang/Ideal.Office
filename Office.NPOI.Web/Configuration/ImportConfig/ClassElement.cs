namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;
    public class ClassElement:ConfigurationElement
    {
        #region 字段
        private const string KeyAttribute = "key";
        private const string ValueAttribute = "value";
        #endregion

        #region 属性

        [ConfigurationProperty(KeyAttribute, IsRequired = true, IsKey = true)]
        public string Key
        {
            get { return (string)this[KeyAttribute]; }
            set { this[KeyAttribute] = value; }
        }

        [ConfigurationProperty(ValueAttribute, IsRequired = false)]
        public string Value
        {
            get { return (string)this[ValueAttribute]; }
            set { this[ValueAttribute] = value; }
        }
        
        #endregion


        #region 方法

        #endregion
    }
}
