using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTemplate.Helper
{
    public static class ObjectHelper
    {
        public static void SetObjectValue(object obj, string fieldPath, object val)
        {
            if (val == null)
            {
                return;
            }

            var currObj = obj;
            var fieldArr = fieldPath.Split('.');

            for (int i = 0; i < fieldArr.Length; i++)
            {
                var props = currObj.GetType().GetProperties();
                var prop = props.FirstOrDefault(a => a.Name == fieldArr[i]);
                if (prop == null)
                {
                    throw new Exception($"类型 {obj.GetType().Name} 内找不到字段 {fieldPath}");
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
                    val = Convert.ChangeType(val, prop.PropertyType);
                    prop.SetValue(currObj, val);
                }

            }
        }

        public static void AddItemToList(object list, object item)
        {
            if (list == null) throw new ArgumentNullException(nameof(list));

            Type listType = list.GetType();

            // 检查是否是List<T>
            if (!listType.IsGenericType || listType.GetGenericTypeDefinition() != typeof(List<>))
            {
                throw new ArgumentException("对象不是泛型List<>");
            }

            // 获取元素类型
            Type elementType = listType.GetGenericArguments()[0];

            // 检查item类型是否匹配
            if (item != null && !elementType.IsAssignableFrom(item.GetType()))
            {
                throw new ArgumentException($"无法将类型{item.GetType()}添加到List<{elementType}>");
            }

            // 获取并调用Add方法
            var addMethod = listType.GetMethod("Add");
            addMethod.Invoke(list, new[] { item });
        }
    }
}
