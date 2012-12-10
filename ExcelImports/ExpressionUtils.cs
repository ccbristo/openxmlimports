using System.Linq.Expressions;
using System.Reflection;
using System;

namespace ExcelImports
{
    public static class ExpressionUtils
    {
        public static MemberInfo GetMemberInfo(this Expression exp)
        {
            // TODO [ccb] This needs better validation;
            var lamdaExp = ((LambdaExpression)exp);
            var memberExp = (MemberExpression)lamdaExp.Body;
            return memberExp.Member;
        }
    }
}
