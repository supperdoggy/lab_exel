using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MyExcel
{
    internal static class AddressesHandler
    {
        public static bool HasAnyAddress(string expr)
        {
            List<string> checkAddressLs = GetAddresses(expr);
            return checkAddressLs.Count != 0;
        }

        public static List<string> GetAddresses(string expr)
        {
            const string pattern = "R[0-9]*C[0-9]*"; // RiCi
            Regex itemRegex = new Regex(pattern, RegexOptions.Compiled);
            MatchCollection matches = itemRegex.Matches(expr);

            List<string> referencesLs = matches
                                        .Cast<Match>()
                                        .Select(m => m.Value)
                                        .ToList();
            referencesLs.RemoveAll(str => str == "RC"); // Remove error-type references from list
            return referencesLs;
        }

        public static List<string> GetAddresses(IEnumerable<MathCell> refCells)
        {
            return refCells.Select(mCell => mCell.OwnAddress).ToList();
        }

        public static (int rowIndex, int colIndex) GetIndexes(string expr)
        {
            string[] indexes = Regex.Split(expr, @"\D+");
            try
            {
                // From 1, because the first element of regexp results is ""
                int refRowIndex = int.Parse(indexes[1]);
                int refColIndex = int.Parse(indexes[2]);
                return (rowIndex: refRowIndex, colIndex: refColIndex);
            }
            catch (System.FormatException)
            {
                throw new InvalidReferenceIndexing();
            }
        }
    }
}