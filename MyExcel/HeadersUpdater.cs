using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace MyExcel
{
    // Context
    internal class HeadersUpdater
    {
        private readonly GridForm _ownerGrid;
        private readonly List<IHeadUpdate> _updates;

        public HeadersUpdater(GridForm grid, params IHeadUpdate[] updates)
        {
            _ownerGrid = grid;
            _updates = updates.ToList();
        }

        public void Update()
        {
            foreach (IHeadUpdate update in _updates)
            {
                update.DoUpdate(_ownerGrid);
            }
        }
    }

    // Strategy
    internal interface IHeadUpdate
    {
        void DoUpdate(GridForm ownerGrid);
    }

    // Concrete Strategy A
    internal class UpdateRows : IHeadUpdate
    {
        public void DoUpdate(GridForm ownerGrid)
        {
            for (int i = 0; i < ownerGrid.dataGridView.RowCount; ++i)
            {
                ownerGrid.dataGridView.Rows[i].HeaderCell.Value = $"{i}";
            }
        }
    }

    // Concrete Strategy B
    internal class UpdateColumns : IHeadUpdate
    {
        public void DoUpdate(GridForm ownerGrid)
        {
            for (int j = 0; j < ownerGrid.dataGridView.ColumnCount; ++j)
            {
                ownerGrid.dataGridView.Columns[j].SortMode = DataGridViewColumnSortMode.NotSortable;
                ownerGrid.dataGridView.Columns[j].MinimumWidth = 100;
                ownerGrid.dataGridView.Columns[j].HeaderCell.Value = $"{j}";
            }
        }
    }
}