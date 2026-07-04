using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Exceptions
{
    /// <summary>
    /// 模版设计异常
    /// </summary>
    public class TemplateDesignException : Exception
    {
        /// <summary>
        /// 实例化模版设计异常
        /// </summary>
        /// <param name="errorType">错误类型</param>
        /// <param name="message">异常信息</param>
        public TemplateDesignException(TemplateDesignExceptionType errorType, string message) : base(message)
        {
            this.ErrorType = errorType;
        }

        /// <summary>
        /// 错误类型
        /// </summary>
        public TemplateDesignExceptionType ErrorType { get; private set; }
    }

    /// <summary>
    /// 模版设计异常类型
    /// </summary>
    public enum TemplateDesignExceptionType
    {
        /// <summary>
        /// 字段冲突
        /// </summary>
        FieldConflict,
        /// <summary>
        /// 位置冲突
        /// </summary>
        PositionConflict,
    }
}
