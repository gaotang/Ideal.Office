using System.Data;

namespace Ideal.Office.Data
{
    /// <summary>
    /// 创建<see cref="IDbConnection"/>的实例
    /// </summary>
    /// <returns></returns>
    public delegate IDbConnection ConnectionObjBuild();
}