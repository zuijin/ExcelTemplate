using KellermanSoftware.CompareNetObjects;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelTemplate.Helper
{
    public static class ObjectHelper
    {
        private static readonly ConcurrentDictionary<string, string[]> _pathCache = new ConcurrentDictionary<string, string[]>();
        private static readonly ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>> _propertyCache = new ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>>();

        private static string[] GetPathArrayCached(string fieldPath)
        {
            return _pathCache.GetOrAdd(fieldPath, p => p.Split('.'));
        }

        private static PropertyInfo GetPropertyCached(Type type, string propertyName)
        {
            var props = _propertyCache.GetOrAdd(type, t =>
            {
                var dict = new Dictionary<string, PropertyInfo>();
                foreach (var p in t.GetProperties())
                {
                    dict[p.Name] = p;
                }
                return dict;
            });

            props.TryGetValue(propertyName, out var prop);
            return prop;
        }

        /// <summary>
        /// 设置对象字段值
        /// </summary>
        /// <param name="obj">对象</param>
        /// <param name="fieldPath">字段路径</param>
        /// <param name="val">值</param>
        public static void SetObjectValue(object obj, string fieldPath, object val)
        {
            if (val == null)
            {
                return;
            }

            var currObj = obj;
            var fieldArr = GetPathArrayCached(fieldPath);

            for (int i = 0; i < fieldArr.Length; i++)
            {
                var prop = GetPropertyCached(currObj.GetType(), fieldArr[i]);
                if (prop == null)
                {
                    throw new ArgumentException($"类型 {currObj.GetType().Name} 内找不到字段 {fieldArr[i]}");
                }

                if (!prop.CanWrite)
                {
                    throw new ArgumentException($"字段 {fieldArr[i]} 无法写入，请检查是否处于只读状态");
                }

                if (i < (fieldArr.Length - 1))
                {
                    var tmp = prop.GetValue(currObj);
                    if (tmp == null)
                    {
                        tmp = Activator.CreateInstance(prop.PropertyType);
                        prop.SetValue(currObj, tmp);
                    }

                    currObj = tmp;
                }
                else
                {
                    // 处理可空类型 (Nullable<T>)
                    var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                    var convertedVal = Convert.ChangeType(val, targetType);
                    prop.SetValue(currObj, convertedVal);
                }
            }
        }

        /// <summary>
        /// 获取对象字段值
        /// </summary>
        /// <param name="obj">对象</param>
        /// <param name="fieldPath">字段路径</param>
        /// <returns>字段值</returns>
        public static object GetObjectValue(object obj, string fieldPath)
        {
            var currObj = obj;
            var fieldArr = GetPathArrayCached(fieldPath);

            for (int i = 0; i < fieldArr.Length; i++)
            {
                var prop = GetPropertyCached(currObj.GetType(), fieldArr[i]);
                if (prop == null)
                {
                    throw new ArgumentException($"类型 {currObj.GetType().Name} 内找不到字段 {fieldArr[i]}");
                }

                if (i < (fieldArr.Length - 1))
                {
                    var tmp = prop.GetValue(currObj);
                    if (tmp == null)
                    {
                        return null;
                    }

                    currObj = tmp;
                }
                else
                {
                    return prop.GetValue(currObj);
                }
            }

            return null;
        }

        /// <summary>
        /// 向列表中添加元素
        /// </summary>
        /// <param name="list">列表对象</param>
        /// <param name="item">元素对象</param>
        public static void AddItemToList(object list, object item)
        {
            if (list == null) throw new ArgumentNullException(nameof(list));

            var listType = list.GetType();

            // 检查是否是List<T>
            if (!listType.IsGenericType || listType.GetGenericTypeDefinition() != typeof(List<>))
            {
                throw new ArgumentException("对象不是泛型List<>");
            }

            // 获取元素类型
            var elementType = listType.GetGenericArguments()[0];

            // 检查item类型是否匹配
            if (item != null && !elementType.IsAssignableFrom(item.GetType()))
            {
                throw new ArgumentException($"无法将类型{item.GetType()}添加到List<{elementType}>");
            }

            // 直接通过 IList 接口添加，避免反射 Invoke 性能损耗
            if (list is System.Collections.IList iList)
            {
                iList.Add(item);
            }
        }

        /// <summary>
        /// 比较两个对象是否相等
        /// </summary>
        /// <param name="obj1">对象1</param>
        /// <param name="obj2">对象2</param>
        /// <returns>是否相等</returns>
        public static bool Compare(object obj1, object obj2)
        {
            var compareLogic = new CompareLogic();
            var result = compareLogic.Compare(obj1, obj2);
            return result.AreEqual;
        }
    }
}
