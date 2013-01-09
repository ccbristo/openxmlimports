using System.Reflection;
using System.Text.RegularExpressions;

namespace OpenXmlImports
{
    public class CamelCaseNamingConvention : INamingConvention
    {
        private static readonly Regex NonLowerFollowedByUpperLower = new Regex(@"(\P{Ll})\B(\P{Ll}\p{Ll})", RegexOptions.Compiled);
        private static readonly Regex LowerCaseFollowedByNonLower = new Regex(@"(\p{Ll})\B(\P{Ll})", RegexOptions.Compiled);

        public string GetName(MemberInfo member)
        {
            string noUnderscores = member.Name.Replace("_", " ");
            string spacesAddedIn_AAb = NonLowerFollowedByUpperLower.Replace(noUnderscores, "$1 $2");
            string spacesAddedIn_aB = LowerCaseFollowedByNonLower.Replace(spacesAddedIn_AAb, "$1 $2");
            return spacesAddedIn_aB;
        }
    }
}
