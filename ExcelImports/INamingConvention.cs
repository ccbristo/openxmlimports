using System.Reflection;

namespace ExcelImports
{
    public interface INamingConvention
    {
        string GetName(MemberInfo member);
    }
}
