using System.Collections.Generic;
using System.Windows.Forms;

namespace MyExcel
{
    internal class MathCellsProvider
    {
        private MathCellsProvider() { }

        private static MathCellsProvider _instance;
        public static MathCellsProvider GetInstance => _instance ?? (_instance = new MathCellsProvider());

        private List<MathCell> UsedCells { get; } = new List<MathCell>();

        public GridForm OwnerGrid { get; private set; }
        public void SetGridForm(GridForm initGrid)
        {
            OwnerGrid = initGrid;
        }

        public MathCell GetMathCell(string address)
        {
            var (rowIndex, colIndex) = AddressesHandler.GetIndexes(address);
            MathCell mCell = (MathCell)OwnerGrid.dataGridView[colIndex, rowIndex].Tag;
            return mCell;
        }
        public static MathCell GetMathCell(DataGridViewCell currCell)
        {
            MathCell mCell = (MathCell)currCell.Tag;
            return mCell;
        }

        public void AddCell(MathCell mCell)
        {
            UsedCells.Add(mCell);
        }

        public void RemoveCell(MathCell mCell)
        {
            UsedCells.Remove(mCell);
        }

        public List<MathCell> GetUsedCells()
        {
            return UsedCells;
        }

        public void ResetMathCells()
        {
            UsedCells.Clear();
        }

        public MathCell ProcessCellInput(DataGridViewCell inputCell)
        {
            string newFormula = string.Empty;
            if (inputCell.Value != null)
            {
                newFormula = inputCell.Value.ToString();
            }

            MathCell mCell = GetMathCell(inputCell);
            mCell.Formula = newFormula;
            mCell.RowIndex = inputCell.RowIndex;
            mCell.ColumnIndex = inputCell.ColumnIndex;

            object prevValue = mCell.Value;
            mCell.EvaluateFormula();
            object currentValue = mCell.Value;

            if (prevValue != null && prevValue.ToString() == "RecursiveReference" ||
                currentValue.ToString() == "RecursiveReference")
            {
                foreach (MathCell usedCell in UsedCells)
                {
                    usedCell.RemoveFromBoundedCells();
                    OnGridMathCellsProvider.UpdateValuesOnGrid(UsedCells, OwnerGrid);
                }
            }

            OnGridMathCellsProvider.UpdateDependentsOnGrid(mCell, OwnerGrid);

            return mCell;
        }
    }
}