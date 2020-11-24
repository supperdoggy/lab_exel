using System.Linq;
using System.Windows.Forms;

namespace MyExcel
{
    internal static class GridModifier
    {
        public static void AddRow(GridForm grid)
        {
            ++grid.dataGridView.RowCount;

            OnGridMathCellsProvider.InitEmptyCells(grid);

            HeadersUpdater headersUpdater = new HeadersUpdater(grid, new UpdateRows());
            headersUpdater.Update();
        }

        public static void AddColumn(GridForm grid)
        {
            ++grid.dataGridView.ColumnCount;

            OnGridMathCellsProvider.InitEmptyCells(grid);

            HeadersUpdater headersUpdater = new HeadersUpdater(grid, new UpdateColumns());
            headersUpdater.Update();
        }

        public static void InsertRow(GridForm grid)
        {
            int rowIndex = grid.dataGridView.CurrentCell.RowIndex;

            grid.dataGridView.Rows.Insert(rowIndex);

            OnGridMathCellsProvider.InitEmptyCells(grid);

            MathCellsShifter mathCellsShifter = new MathCellsShifter(new ShiftAfterRowInsert(), rowIndex, grid);
            mathCellsShifter.Shift();

            HeadersUpdater headersUpdater = new HeadersUpdater(grid, new UpdateRows());
            headersUpdater.Update();
        }

        public static void InsertColumn(GridForm grid)
        {
            int colIndex = grid.dataGridView.CurrentCell.ColumnIndex;

            DataGridViewTextBoxColumn dummyCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "dummyColText"
            };

            grid.dataGridView.Columns.Insert(colIndex, dummyCol);

            OnGridMathCellsProvider.InitEmptyCells(grid);

            MathCellsShifter mathCellsShifter = new MathCellsShifter(new ShiftAfterColumnInsert(), colIndex, grid);
            mathCellsShifter.Shift();

            HeadersUpdater headersUpdater = new HeadersUpdater(grid, new UpdateColumns());
            headersUpdater.Update();
        }

        public static void DeleteRow(GridForm grid)
        {
            if (grid.dataGridView.RowCount <= 1) { return; }

            int rowIndex = grid.dataGridView.CurrentCell.RowIndex;

            bool isDeletionAllowed = true;
            if (IsRowHasValue(grid, rowIndex))
            {
                isDeletionAllowed = IsDeletionAllowed("row");
            }

            if (isDeletionAllowed)
            {
                DataGridViewRow deletionRow = grid.dataGridView.Rows[rowIndex];

                foreach (DataGridViewCell cell in deletionRow.Cells)
                {
                    MathCell delCell = MathCellsProvider.GetMathCell(cell);
                    delCell.UpdateDependentsBeforeDelete();
                    MathCellsProvider.GetInstance.RemoveCell(delCell);
                }

                grid.dataGridView.Rows.RemoveAt(rowIndex);

                MathCellsShifter mathCellsShifter = new MathCellsShifter(new ShiftAfterRowDeletion(), rowIndex, grid);
                mathCellsShifter.Shift();

                HeadersUpdater headersUpdater = new HeadersUpdater(grid, new UpdateRows());
                headersUpdater.Update();
            }
        }

        public static void DeleteColumn(GridForm grid)
        {
            if (grid.dataGridView.ColumnCount <= 1) { return; }

            int colIndex = grid.dataGridView.CurrentCell.ColumnIndex;

            bool isDeletionAllowed = true;
            if (IsColumnHasValue(grid, colIndex))
            {
                isDeletionAllowed = IsDeletionAllowed("column");
            }

            if (isDeletionAllowed)
            {
                foreach (DataGridViewRow row in grid.dataGridView.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.ColumnIndex == colIndex)
                        {
                            MathCell delCell = MathCellsProvider.GetMathCell(cell);
                            delCell.UpdateDependentsBeforeDelete();
                            MathCellsProvider.GetInstance.RemoveCell(delCell);
                        }
                    }
                }

                grid.dataGridView.Columns.RemoveAt(colIndex);

                MathCellsShifter mathCellsShifter = new MathCellsShifter(new ShiftAfterColumnDeletion(), colIndex, grid);
                mathCellsShifter.Shift();

                HeadersUpdater headersUpdater = new HeadersUpdater(grid, new UpdateColumns());
                headersUpdater.Update();
            }
        }

        private static bool IsDeletionAllowed(string deleteItem)
        {
            DialogResult removeDiagResult = MessageBox.Show($"This {deleteItem} has values.\n" +
                                                            $"Do you really want to delete this {deleteItem}?",
                                                            $"Delete {deleteItem}", MessageBoxButtons.YesNo);

            return removeDiagResult != DialogResult.No;
        }

        private static bool IsRowHasValue(GridForm grid, int rowIndex)
        {
            DataGridViewRow deletionRow = grid.dataGridView.Rows[rowIndex];
            var hasValue = (from DataGridViewCell cell in deletionRow.Cells
                            select cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                            .Any(r => r);
            return hasValue;
        }

        private static bool IsColumnHasValue(in GridForm grid, int colIndex)
        {
            bool hasValue = (from DataGridViewRow row in grid.dataGridView.Rows
                             from DataGridViewCell cell in row.Cells
                             where cell.ColumnIndex == colIndex
                             select cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                             .Any(c => c);
            return hasValue;
        }
    }
}