namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    internal class CatalogElement : ConfigurationElement
    {
        #region 字段
        private const string KeyAttribute = "key";
        private const string TemplateAttribute = "template";
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

        [ConfigurationProperty(TemplateAttribute, IsDefaultCollection = true)]
        public TemplateElementCollection Template
        {
            get
            {
                return (TemplateElementCollection)base[TemplateAttribute];
            }
        }
        #endregion
    }
}
