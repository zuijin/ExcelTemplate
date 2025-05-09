using System;
using System.Linq.Expressions;
using ExcelTemplate.Model;

namespace ExcelTemplate.Hint
{
    public class FieldHintExp<T, T2>
    {
        private HintBuilder<T> _builder;
        private Expression _expression;


        public FieldHintExp(HintBuilder<T> builder, Expression expression)
        {
            _builder = builder;
            _expression = expression;
        }

        public void AddError(string message)
        {
            var pos = GetPosition();
            _builder.AddError(pos.Row, pos.Col, message);
        }

        private Position GetPosition()
        {
            var dataPath = Visit(_expression);
            return _builder.FieldPositionDic[dataPath];
        }

        private string Visit(Expression exp)
        {
            if (exp == null)
            {
                return string.Empty;
            }

            switch (exp.NodeType)
            {
                case ExpressionType.Lambda:
                    return VisiLambda((LambdaExpression)exp);
                case ExpressionType.MemberAccess:
                    return VisitMemberAccess((MemberExpression)exp);

                case ExpressionType.Call:
                    return VisitMethodCall((MethodCallExpression)exp);
                default:
                    throw new Exception("不支持的表达式");

            }
        }

        private string VisiLambda(LambdaExpression exp)
        {
            return Visit(exp.Body);
        }

        private string VisitMemberAccess(MemberExpression exp)
        {
            if (exp.Expression.NodeType == ExpressionType.Parameter)
            {
                return exp.Member.Name;
            }

            var preStr = Visit(exp.Expression);
            return $"{preStr}.{exp.Member.Name}";
        }

        private string VisitMethodCall(MethodCallExpression exp)
        {
            if (exp.Method.ReflectedType.FullName == "ExcelTemplate.Hint.HintBuilderExtensions" && exp.Method.Name == "Pick")
            {
                var preStr = Visit(exp.Arguments[0]);
                var index = -1;
                var arg = exp.Arguments[1];
                if (arg.NodeType == ExpressionType.MemberAccess)
                {
                    var lambda = Expression.Lambda(arg).Compile();
                    var obj = lambda.DynamicInvoke();
                    if (obj is int)
                    {
                        index = (int)obj;
                    }
                    else
                    {
                        if (!_builder.ElemetIndexDic.ContainsKey(obj))
                        {
                            throw new Exception("该对象不属于集合内");
                        }

                        index = _builder.ElemetIndexDic[obj];
                    }
                }
                else if (arg.NodeType == ExpressionType.Constant)
                {
                    index = (int)((ConstantExpression)arg).Value;
                }
                else
                {
                    throw new Exception("不支持的参数表达式");
                }

                return $"{preStr}.{index}";
            }

            throw new Exception($"不支持的表达式{exp.Method.Name}");
        }
    }
}
