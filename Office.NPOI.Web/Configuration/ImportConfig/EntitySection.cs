namespace Ideal.Office.Web.Configuration
{
    using System.Configuration;

    internal class EntitySection: ConfigurationSection
    {
        #region 字段
        private const string SectionPath = "excel/importconfig";
        private const string EntityElement = "";
        #endregion

        #region 属性
        /// <summary>
        /// 获取集合中的目录配置.
        /// </summary>
        [ConfigurationProperty(EntityElement, IsDefaultCollection = true)]
        public EntityElementCollection Entity
        {
            get
            {
                return (EntityElementCollection)base[EntityElement];
            }
        }
        #endregion

        #region 方法

        /// <summary>
        /// 在当前配置中获取实例 <see cref="CompositionConfigurationSection" />.
        /// </summary>
        /// <returns>返回实例 <see cref="CompositionConfigurationSection" />, 或NULL实例.</returns>
        public static EntitySection GetInstance()
        {
            return ConfigurationManager.GetSection(SectionPath) as EntitySection;
        }
        #endregion
    }
}
