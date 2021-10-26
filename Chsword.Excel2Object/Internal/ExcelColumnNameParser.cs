using System.Collections.Generic;

namespace Chsword.Excel2Object.Internal
{
    internal static class ExcelColumnNameParser
    {
        public static string Parse(int columnIndex, int initValue = 0)
        {
            var x = columnIndex;
            var list = new List<char>();

            do
            {
                var mod = x % 26;
                if (x != columnIndex)
                {
                    mod -= 1;
                }

                x /= 26;
                list.Add((char)('A' + mod));
            } while (x > 0);

            list.Reverse();
            return string.Join("", list);
        }
    }
}