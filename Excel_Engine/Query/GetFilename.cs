using System.Text.RegularExpressions;

namespace BH.Engine.Excel
{
    public static partial class Query
    {
        public static string GetFilename()
        {
            string reftext = Caller().RefText();
            return Regex.Match(reftext, @"\[(.*)\]").Groups[1].Value;
        }
    }
}