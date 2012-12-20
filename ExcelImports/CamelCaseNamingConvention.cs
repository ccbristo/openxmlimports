using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelImports
{
    public class CamelCaseNamingConvention : INamingConvention
    {
        private static readonly Regex UpperCaseFollowedByUpperLower = new Regex(@"(\P{Ll})(\P{Ll}\p{Ll})", RegexOptions.Compiled);
        private static readonly Regex LowerCaseFollowedByNonLower = new Regex(@"(\p{Ll})(\P{Ll})", RegexOptions.Compiled);

        public string GetName(MemberInfo member)
        {
            string r1 = UpperCaseFollowedByUpperLower.Replace(member.Name, "$1 $2");
            string r2 = LowerCaseFollowedByNonLower.Replace(r1, "$1 $2");
            return r2;
        }
    }
}
