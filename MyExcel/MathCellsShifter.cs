using System.Collections.Generic;
using System.Windows.Forms;

namespace MyExcel
{
    // Context
    internal class MathCellsShifter
    {
        private readonly IShift _shiftWay;
        private readonly int _delimiterIndex;
        private readonly GridForm _ownerGrid;

        public MathCellsShifter(IShift shiftWay, int delimiterIndex, GridForm ownerGrid)
        {
            _shiftWay = shiftWay;
            _delimiterIndex = delimiterIndex;
            _ownerGrid = ownerGrid;
        }

        public void Shift()
        {
            foreach (DataGridViewRow row in _ownerGrid.dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    MathCell mCell = MathCellsProvider.GetMathCell(cell);
                    _shiftWay.DoShift(_delimiterIndex, mCell, _ownerGrid);
                }
            }

            List<MathCell> usedMathCells = MathCellsProvider.GetInstance.GetUsedCells();
            foreach (MathCell usedCell in usedMathCells)
            {
                usedCell.ShiftReferences();
            }

            OnGridMathCellsProvider.UpdateValuesOnGrid(usedMathCells, _ownerGrid);
        }
    }

    // Strategy
    internal interface IShift
    {
        void DoShift(int delimiterIndex, MathCell mCell, GridForm ownerGrid);
    }

    // Concrete Strategy A
    internal class ShiftAfterRowDeletion : IShift // To up after certain row
    {
        public void DoShift(int delimiterIndex, MathCell mCell, GridForm ownerGrid)
        {
            int rowIndex = mCell.RowIndex;
            if (rowIndex > 0 && rowIndex > delimiterIndex && rowIndex <= ownerGrid.dataGridView.RowCount)
            {
                --mCell.RowIndex;
            }
        }
    }

    // Concrete Strategy B
    internal class ShiftAfterColumnDeletion : IShift // To left after certain column
    {
        public void DoShift(int delimiterIndex, MathCell mCell, GridForm ownerGrid)
        {
            int colIndex = mCell.ColumnIndex;
            if (colIndex > 0 && colIndex > delimiterIndex && colIndex <= ownerGrid.dataGridView.ColumnCount)
            {
                --mCell.ColumnIndex;
            }
        }
    }

    // Concrete Strategy C
    internal class ShiftAfterRowInsert : IShift // To right
    {
        public void DoShift(int delimiterIndex, MathCell mCell, GridForm ownerGrid)
        {
            int rIndex = mCell.RowIndex;
            if (rIndex >= delimiterIndex && rIndex < ownerGrid.dataGridView.RowCount)
            {
                ++mCell.RowIndex;
            }
        }
    }

    // Concrete Strategy D
    internal class ShiftAfterColumnInsert : IShift // To down
    {
        public void DoShift(int delimiterIndex, MathCell mCell, GridForm ownerGrid)
        {
            int cIndex = mCell.ColumnIndex;
            if (cIndex >= delimiterIndex && cIndex < ownerGrid.dataGridView.ColumnCount)
            {
                ++mCell.ColumnIndex;
            }
        }
    }
}