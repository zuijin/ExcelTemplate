using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Exceptions
{
    public class TemplateDesignException : Exception
    {
        public TemplateDesignException(TemplateDesignExceptionType errorType, string message) : base(message)
        {
            this.ErrorType = errorType;
        }

        public TemplateDesignExceptionType ErrorType { get; private set; }
    }

    public enum TemplateDesignExceptionType
    {
        FieldConflict,
        PositionConflict,
    }
}
