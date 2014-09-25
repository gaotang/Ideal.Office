/**
 * Haibin Zou=>zhb
 * hibean2006@126.com
 * 2011-11-18 Created 
 * */
using System;

namespace Ideal.Office.Data
{
    /// <summary>
    /// 数据库连接字符串未能找到的异常
    /// </summary>
    public class DbConnectionNotFoundException : ApplicationException
    {
        /// <summary>
        /// 初始化<see cref="DbConnectionNotFoundException"/>
        /// </summary>
        public DbConnectionNotFoundException()
            : base("没有在 config（web.config/app.config/machine.config） 文件中找到连接字符串信息")
        {
            
        }

        /// <summary>
        /// 初始化<see cref="DbConnectionNotFoundException"/>的实例
        /// </summary>
        public DbConnectionNotFoundException(string name)
            : base(string.Format("没有在 config（web.config/app.config/machine.config） 文件中找到连接字符串信息, 名称为：{0}", name))
        {

        }
    }
}