using System.Reflection;

namespace OpenXmlImports
{
    public interface INamingConvention
    {
        string GetName(MemberInfo member);
    }
}
