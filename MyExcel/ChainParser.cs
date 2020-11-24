using org.mariuszgromada.math.mxparser;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MyExcel
{
    interface IParser
    {
        IParser SetNextParser(IParser parseAlgo);
        object Parse(MathCell ownerMathCell);
    }

    abstract class FormulaParser : IParser
    {
        private IParser _nextParseAlgo;
        public IParser SetNextParser(IParser parseAlgo)
        {
            _nextParseAlgo = parseAlgo;
            return parseAlgo; // to allow chain setting of algorithm
        }

        public virtual object Parse(MathCell ownerMathCell)
        {
            return _nextParseAlgo?.Parse(ownerMathCell);
        }
    }

    class IsExpressionParser : FormulaParser
    {
        public override object Parse(MathCell ownerMathCell)
        {
            // Removes spaces for proper parsing
            string formula = ownerMathCell.Formula = RemoveSpaces(ownerMathCell.Formula);

            List<string> refAddresses = AddressesHandler.GetAddresses(formula);
            if (refAddresses.Count != 0 && formula.First() != '=')
            {
                ownerMathCell.ResetReferences();
                throw new InvalidReferenceFormat();
            }

            if (string.IsNullOrEmpty(formula) || formula.First() != '=')
            {
                ownerMathCell.ResetReferences();
                return formula;
            }

            return base.Parse(ownerMathCell); // Proceed
        }

        private static string RemoveSpaces(string input)
        {
            return Regex.Replace(input, @"\s+", "", RegexOptions.Compiled);
        }
    }

    class ReferenceFormattingParser : FormulaParser
    {
        public override object Parse(MathCell ownerMathCell)
        {
            string formula = ownerMathCell.Formula;
            if (HasAnyRefSign(formula))
            {
                if (HasInvalidRefFormat(formula))
                {
                    throw new InvalidReferenceFormat();
                }

                if (HasInvalidIndexing(formula))
                {
                    throw new InvalidReferenceIndexing();
                }
            }

            return base.Parse(ownerMathCell);
        }

        private static bool HasAnyRefSign(string formula)
        {
            return formula.Contains('R') || formula.Contains('C') ||
                   formula.Contains('r') || formula.Contains('c');
        }

        private static bool HasInvalidRefFormat(string formula)
        {
            bool hasInvalidRefFormat = formula.Contains("RC") || // Error state reference
                                       formula.Contains('R') && !formula.Contains('C') || // Contains R without C
                                       formula.Contains('C') && !formula.Contains('R') || // Contains C without R
                                       formula.Contains('r') && formula.Contains('c') ||
                                       formula.IndexOf('C') < formula.IndexOf('R'); //Reference given in CiRi notation
            return hasInvalidRefFormat;
        }

        private static bool HasInvalidIndexing(string formula)
        {
            List<string> formulaAddresses = AddressesHandler.GetAddresses(formula);
            GridForm ownerGrid = MathCellsProvider.GetInstance.OwnerGrid;
            foreach (string address in formulaAddresses)
            {
                var (rowIndex, colIndex) = AddressesHandler.GetIndexes(address);
                if (rowIndex >= ownerGrid.dataGridView.RowCount || colIndex >= ownerGrid.dataGridView.ColumnCount)
                {
                    return true;
                }
            }
            return false;
        }
    }

    class DependenciesParser : FormulaParser
    {
        public override object Parse(MathCell ownerMathCell)
        {
            ownerMathCell.ResetReferences();

            string formula = ownerMathCell.Formula;
            if (AddressesHandler.HasAnyAddress(formula))
            {
                List<string> formulaReferences = AddressesHandler.GetAddresses(formula);
                for (int i = 0; i < formulaReferences.Count; ++i)
                {
                    string reference = formulaReferences[i];

                    MathCell refCell =
                        MathCellsProvider.GetInstance.GetMathCell(reference);

                    if (!ownerMathCell.References.Contains(refCell))
                    {
                        ownerMathCell.InsertReference(refCell);
                    }

                    if (!refCell.Dependents.Contains(ownerMathCell))
                    {
                        refCell.InsertDependent(ownerMathCell);
                    }
                }
            }

            return base.Parse(ownerMathCell);
        }
    }

    class RecursionParser : FormulaParser
    {
        public override object Parse(MathCell ownerMathCell)
        {
            List<MathCell> refAddressesCheckList = new List<MathCell> { ownerMathCell };
            RecursionCheck(ownerMathCell, refAddressesCheckList);

            return base.Parse(ownerMathCell);
        }

        private static void RecursionCheck(MathCell refCell, List<MathCell> refCheckList)
        {
            List<MathCell> inspectRef = refCell.References;

            List<MathCell> recursiveCells = refCheckList.AsQueryable().Intersect(inspectRef).ToList();
            if (recursiveCells.Count != 0)
            {
                foreach (MathCell recursiveCell in recursiveCells)
                {
                    recursiveCell.RemoveFromReferences();
                }
                throw new RecursiveReference();
            }

            foreach (MathCell depCell in inspectRef)
            {
                refCheckList.Add(depCell);
                RecursionCheck(depCell, refCheckList);
            }
        }
    }

    class Dereference : FormulaParser
    {
        public override object Parse(MathCell ownerMathCell)
        {
            string formula = ownerMathCell.DereferencedFormula =
                ownerMathCell.Formula = RemoveRedundantZeroes(ownerMathCell);

            if (AddressesHandler.HasAnyAddress(formula))
            {
                int i = 0;
                while (AddressesHandler.HasAnyAddress(formula)) // While formula contains any reference
                {
                    MathCell refCell = ownerMathCell.References[i];
                    string refAddress = refCell.OwnAddress;

                    if (refCell.Value == null)
                    {
                        refCell.EvaluateFormula();
                    }

                    string refValue = refCell.Value.ToString();
                    if (string.IsNullOrEmpty(refValue))
                    {
                        ownerMathCell.DereferencedFormula = formula = formula.Replace(refAddress, "0");
                    }
                    ownerMathCell.DereferencedFormula = formula = formula.Replace(refAddress, refValue);

                    ++i;
                }
            }

            return base.Parse(ownerMathCell);
        }

        private static string RemoveRedundantZeroes(MathCell ownerMathCell)
        {
            string formula = ownerMathCell.Formula;
            List<string> oldAdr = AddressesHandler.GetAddresses(formula);

            const string pattern = @"0+(?=(0|[1-9]))";
            for (int i = 0; i < oldAdr.Count; ++i)
            {
                string updStr = Regex.Replace(oldAdr[i], pattern, "", RegexOptions.Compiled);
                formula = ownerMathCell.Formula.Replace(oldAdr[i], updStr);
            }
            return formula;
        }
    }

    class ExpressionParser : FormulaParser
    {
        public override object Parse(MathCell ownerMathCell)
        {
            string formula = ownerMathCell.DereferencedFormula;

            if (HasDivByZero(formula))
            {
                throw new DivByZero();
            }

            formula = ownerMathCell.DereferencedFormula = SurroundNegativeWithBrackets(formula);
            formula = ownerMathCell.DereferencedFormula = formula.Substring(formula.IndexOf('=') + 1); // Removes =

            Expression expr = new Expression(formula);
            bool isValidOperation = expr.checkSyntax(); // mXParser built-in method
            if (!isValidOperation)
            {
                throw new InvalidOperation();
            }

            return base.Parse(ownerMathCell);
        }

        private static string SurroundNegativeWithBrackets(string expr)
        {
            const string splitPattern = @"[\+\*\-\/\^]?(-?\d+)";
            string[] terms = Regex.Split(expr, splitPattern, RegexOptions.Compiled);

            const string negativeValuesPattern = @"-((0\.\d*[1-9])|([1-9]\d*)(\.\d+)?)";
            foreach (var term in terms)
            {
                Regex itemRegex = new Regex(negativeValuesPattern, RegexOptions.Compiled);
                MatchCollection match = itemRegex.Matches(term);
                if (match.Count > 0)
                {
                    string negativeNumber = match[0].Value;
                    expr = expr.Replace(negativeNumber, $"({negativeNumber})");
                }
            }

            return expr;
        }

        private static bool HasDivByZero(string input)
        {
            const string pattern = @".*\/0([^.]|$|\.(0{4,}.*|0{1,4}([^0-9]|$))).*";
            Regex itemRegex = new Regex(pattern, RegexOptions.Compiled);
            MatchCollection match = itemRegex.Matches(input);
            return match.Count != 0;
        }
    }

    class Calculate : FormulaParser
    {
        public override object Parse(MathCell ownerMathCell)
        {
            string formula = ownerMathCell.DereferencedFormula;
            Expression expr = new Expression(formula);
            return expr.calculate();
        }
    }
}
