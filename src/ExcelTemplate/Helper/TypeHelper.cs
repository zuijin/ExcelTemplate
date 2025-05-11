using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTemplate.Helper
{
    public class TypeHelper
    {
        /// <summary>
        /// 判断是否简单类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsSimpleType(Type type)
        {
            if (type == typeof(string) || type == typeof(DateTime))
                return true;

            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                return IsSimpleType(type.GetGenericArguments()[0]);

            return (type.IsValueType && type.IsPrimitive) || type.IsEnum;
        }

        /// <summary>
        /// 判断是否集合类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsCollectionType(Type type)
        {
            if (type == typeof(string))
                return false;

            // 处理Nullable类型
            var underlyingType = Nullable.GetUnderlyingType(type);
            if (underlyingType != null)
                type = underlyingType;

            // 数组
            if (type.IsArray)
                return true;

            // 检查IEnumerable<>接口
            if (type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>)))
            {
                // 排除字典类型
                if (typeof(System.Collections.IDictionary).IsAssignableFrom(type))
                    return false;

                return true;
            }

            // 非泛型集合
            return typeof(System.Collections.IEnumerable).IsAssignableFrom(type);
        }

        /// <summary>
        /// 获取集合元素的类型
        /// </summary>
        /// <param name="collectionType"></param>
        /// <returns></returns>
        public static Type GetCollectionElementType(Type collectionType)
        {
            // 处理数组
            if (collectionType.IsArray)
                return collectionType.GetElementType();

            // 处理IEnumerable<T>
            if (collectionType.IsGenericType &&
                collectionType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            {
                return collectionType.GetGenericArguments()[0];
            }

            // 处理实现了IEnumerable<T>的类型
            var enumerableInterface = collectionType.GetInterfaces()
                .FirstOrDefault(i => i.IsGenericType &&
                                   i.GetGenericTypeDefinition() == typeof(IEnumerable<>));

            if (enumerableInterface != null)
            {
                return enumerableInterface.GetGenericArguments()[0];
            }

            // 非泛型集合（如ArrayList）返回object
            return typeof(object);
        }

        /// <summary>
        /// 检查是否派生自泛型类
        /// </summary>
        /// <param name="genericBaseType"></param>
        /// <param name="typeToCheck"></param>
        /// <returns></returns>
        public static bool IsSubclassOfRawGeneric(Type genericBaseType, Type typeToCheck)
        {
            if (genericBaseType == null || typeToCheck == null)
                return false;

            // 检查当前类型
            if (typeToCheck.IsGenericType && typeToCheck.GetGenericTypeDefinition() == genericBaseType)
            {
                return true;
            }

            // 检查基类
            if (typeToCheck.BaseType != null && IsSubclassOfRawGeneric(genericBaseType, typeToCheck.BaseType))
            {
                return true;
            }

            // 检查实现的接口
            return typeToCheck.GetInterfaces().Any(interfaceType => interfaceType.IsGenericType && interfaceType.GetGenericTypeDefinition() == genericBaseType);
        }
    }
}
