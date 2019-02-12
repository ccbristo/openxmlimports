using System.Reflection;
using OpenXmlImports.Core;

namespace OpenXmlImports
{
    public class CamelCaseNamingConvention : INamingConvention
    {
        public string GetName(MemberInfo member)
        {
            return CamelCaseFormatter.Format(member.Name);
        }
    }
}