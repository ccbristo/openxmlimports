using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace OpenXmlImports
{
    public static class ExpressionUtils
    {
        public static MemberInfo GetMemberInfo(this Expression exp)
        {
            // TODO [ccb] This needs better validation;
            var lamdaExp = ((LambdaExpression)exp);
            var memberExp = (MemberExpression)lamdaExp.Body;

            // this would seem like the obvious thing to do, but
            // it doesn't play well with inheritance hierarchies
            //return memberExp.Member;

            var member = memberExp.Expression.Type
                .GetMember(memberExp.Member.Name, MemberTypes.Property | MemberTypes.Field, BindingFlags.Public | BindingFlags.Instance)
                .Single();

            return member;
        }
    }
}
