namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    internal class AddItemElement : ConfigurationElement
    {
        #region 字段
        private const string KeyAttribute = "key";
        private const string ValueAttribute = "value";
        
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

        

        #endregion
    }
}
