namespace OpenXml.Excel.Data
{
    internal static class Error
    {
        public static string NotFoundSheetIndex(int index)
        {
            return string.Format("The sheet with index {0} not found.", index);
        }

        public static string NotFoundSheetName(string name)
        {
            return string.Format("The sheet with name {0} not found.", name);
        } 

    }
}