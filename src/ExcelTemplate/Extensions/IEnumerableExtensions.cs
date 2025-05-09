using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTemplate.Extensions
{
    public static class IEnumerableExtensions
    {
        /// <summary>
        /// 选中集合中的一个元素
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T Pick<T>(this IEnumerable<T> list, T obj)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 选中集合中的一个元素
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T Pick<T>(this IEnumerable<T> list, int index)
        {
            throw new NotImplementedException();
        }
    }
}
