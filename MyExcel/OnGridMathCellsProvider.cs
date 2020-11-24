using System.Collections.Generic;
using System.Windows.Forms;

namespace MyExcel
{
    static class OnGridMathCellsProvider
    {
        public static void InitEmptyCells(GridForm ownerGrid)
        {
            foreach (DataGridViewRow row in ownerGrid.dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Tag == null)
                    {
                        cell.Tag = new MathCell("")
                        {
                            RowIndex = cell.RowIndex,
                            ColumnIndex = cell.ColumnIndex
                        };
                    }
                }
            }
        }

        public static void UpdateDependentsOnGrid(MathCell mCell, GridForm ownerGrid)
        {
            for (int i = 0; i < mCell.Dependents.Count; ++i)
            {
                MathCell dependentCell = mCell.Dependents[i];
                dependentCell.EvaluateFormula();
                PutMathCellOnGrid(dependentCell, ownerGrid);
                UpdateDependentsOnGrid(dependentCell, ownerGrid);
            }
        }

        public static void UpdateValuesOnGrid(List<MathCell> mCells, GridForm ownerGrid)
        {
            foreach (MathCell mCell in mCells)
            {
                mCell.EvaluateFormula();
                PutMathCellOnGrid(mCell, ownerGrid);
            }
        }

        public static void PutMathCellOnGrid(MathCell mCell, GridForm ownerGrid)
        {
            ownerGrid.dataGridView[mCell.ColumnIndex, mCell.RowIndex].Value = mCell.Value;
        }
    }
}