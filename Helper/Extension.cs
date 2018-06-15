using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Helper
{
    /// <summary>
    /// 扩展类
    /// </summary>
    public static class Extension
    {
        /// <summary>
        /// 检查List是不是为空
        /// </summary>
        /// <typeparam name="T">所有继承IEnumerable的类型</typeparam>
        /// <param name="list">要检查的List</param>
        /// <returns>如果为空则返回的List，否则返回原来的List</returns>
        public static IEnumerable<T> CheckNull<T>(this IEnumerable<T> list)
        {
            return list==null?new List<T>(0):list;
        }
    }
}
