namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    internal class CatalogsSection : ConfigurationSection
    {
        #region 字段
        private const string SectionPath = "excel/exportconfig";
        private const string CatalogsElement = "catalogs";
        #endregion

        #region 属性
        /// <summary>
        /// 获取集合中的目录配置.
        /// </summary>
        [ConfigurationProperty(CatalogsElement, IsDefaultCollection = true)]
        public CatalogElementCollection Catalogs
        {
            get
            {
                return (CatalogElementCollection)base[CatalogsElement];
            }
        }
        #endregion

        #region 方法

        /// <summary>
        /// 在当前配置中获取实例 <see cref="CompositionConfigurationSection" />.
        /// </summary>
        /// <returns>返回实例 <see cref="CompositionConfigurationSection" />, 或NULL实例.</returns>
        public static CatalogsSection GetInstance()
        {
            return ConfigurationManager.GetSection(SectionPath) as CatalogsSection;
        }
        #endregion


    }
}
