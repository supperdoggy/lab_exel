using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MyExcel
{
    public class MathCell : DataGridViewCell
    {
        public readonly List<MathCell> References = new List<MathCell>();
        public readonly List<MathCell> Dependents = new List<MathCell>();

        public new int RowIndex { get; set; }
        public new int ColumnIndex { get; set; }

        public string OwnAddress => $"R{RowIndex}C{ColumnIndex}";

        public string Formula { get; set; }
        public string DereferencedFormula { get; set; }

        public new object Value { get; set; }

        public MathCell(string formula)
        {
            Formula = formula;
        }

        public void EvaluateFormula()
        {
            try
            {
                var parserFilters = new IsExpressionParser();
                parserFilters.SetNextParser(new ReferenceFormattingParser())
                    .SetNextParser(new DependenciesParser())
                    .SetNextParser(new RecursionParser())
                    .SetNextParser(new Dereference())
                    .SetNextParser(new ExpressionParser())
                    .SetNextParser(new Calculate());

                Value = parserFilters.Parse(this);
            }
            catch (ParserExceptions pe)
            {
                Value = pe.Message;
            }
        }

        public void InsertReference(MathCell dependCell)
        {
            References.Insert(0, dependCell);
        }

        public void InsertDependent(MathCell dependCell)
        {
            Dependents.Insert(0, dependCell);
        }

        public void ResetReferences()
        {
            References.Clear();
        }

        public void UpdateDependentsBeforeDelete()
        {
            foreach (MathCell dependentCell in Dependents)
            {
                dependentCell.Formula = dependentCell.Formula.Replace(OwnAddress, "RC");
            }
            RemoveFromBoundedCells();
        }

        public void RemoveFromBoundedCells()
        {
            RemoveFromReferences();
            RemoveFromDependents();
        }

        public void RemoveFromReferences()
        {
            foreach (MathCell dependentCell in References)
            {
                dependentCell.Dependents.Remove(this);
            }
        }

        private void RemoveFromDependents()
        {
            foreach (MathCell dependentCell in Dependents)
            {
                dependentCell.References.Remove(this);
            }
        }

        public void ShiftReferences()
        {
            List<string> updatedReferences = AddressesHandler.GetAddresses(References);
            List<string> oldReferences = AddressesHandler.GetAddresses(Formula);

            try
            {
                for (int i = 0; i < oldReferences.Count; ++i)
                {
                    if (oldReferences[i] != "RC" && oldReferences[i] != updatedReferences[i])
                    {
                        Formula = Formula.Replace(oldReferences[i], updatedReferences[i]);
                    }
                }
            }
            catch (ArgumentException ae)
            {
                MessageBox.Show(ae.Message, "Error in references shifting", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}