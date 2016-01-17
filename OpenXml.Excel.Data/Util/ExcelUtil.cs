namespace OpenXml.Excel.Data.Util
{
    public static class ExcelUtil
    {
        public static int GetColumnIndexByName(string colName)
        {
            var name = GetStartingLettersOnly(colName);

            int number = 0, pow = 1;
            for (var i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number - 1;
        }

        private static string GetStartingLettersOnly(string colName)
        {
            var result = string.Empty;
            foreach (var ch in colName)
            {
                if (char.IsLetter(ch))
                    result += ch;
                else
                    break;
            }
            return result;
        }
    }
}