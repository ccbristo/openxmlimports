using System.Reflection;
using System.Text;

namespace ExcelImports
{
    public class CamelCaseNamingConvention : INamingConvention
    {

        public string GetName(MemberInfo member)
        {
            var sb = new StringBuilder(member.Name);

            bool lastWasUpper = char.IsUpper(member.Name[0]);

            for (int i = 2; i < sb.Length - 1; i++)
            {
                char current = sb[i];
                char next = sb[i + 1];

                if (IsUpperToLowerCaseTransition(current, next))
                {
                    sb.Insert(i, " ");
                    i++; // length goes up, keep our pointer to the right
                }
                else if (IsLowerToUpperCaseTransition(current, next) ||
                    IsAlphaToNumericTransition(current, next) ||
                    IsNumericToAlphaTransition(current, next))
                {
                    sb.Insert(i + 1, " ");
                    i += 2; // move to the right of the space we just inserted
                }
            }

            return sb.ToString();
        }

        private static bool IsLowerToUpperCaseTransition(char current, char next)
        {
            return char.IsLower(current) && char.IsUpper(next);
        }

        private static bool IsUpperToLowerCaseTransition(char current, char next)
        {
            return char.IsUpper(current) && char.IsLower(next);
        }

        private static bool IsAlphaToNumericTransition(char current, char next)
        {
            return char.IsLetter(current) && char.IsDigit(next);
        }

        private static bool IsNumericToAlphaTransition(char current, char next)
        {
            return char.IsDigit(current) && char.IsLetter(next);
        }
    }
}
