using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml;

namespace MyExcel
{
    static class FilesHandler
    {
        public static void OpenFile(GridForm grid)
        {
            if (grid.openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = grid.openFileDialog.FileName;

                DataSet dataSet = InstanceDataSet(filePath);
                if (dataSet == null)
                {
                    return;
                }

                MathCellsProvider.GetInstance.ResetMathCells();
                grid.ClearGrid();
                grid.ClearOutputBoxes();

                DataTable dataTable = dataSet.Tables[0];
                grid.dataGridView.ColumnCount = dataTable.Columns.Count;
                grid.dataGridView.RowCount = dataTable.Rows.Count;

                foreach (DataGridViewRow row in grid.dataGridView.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        MathCell newCell = new MathCell(dataTable.Rows[cell.RowIndex][cell.ColumnIndex].ToString())
                        {
                            RowIndex = cell.RowIndex, ColumnIndex = cell.ColumnIndex
                        };

                        if (!string.IsNullOrWhiteSpace(newCell.Formula))
                        {
                            MathCellsProvider.GetInstance.AddCell(newCell);
                        }

                        cell.Tag = newCell;
                    }
                }

                List<MathCell> addedCells = MathCellsProvider.GetInstance.GetUsedCells();
                OnGridMathCellsProvider.UpdateValuesOnGrid(addedCells, grid);
                
                // first value reevalutes because it can be calculated with non proper references
                addedCells.First().EvaluateFormula();
                OnGridMathCellsProvider.PutMathCellOnGrid(addedCells.First(), grid);

                HeadersUpdater headersUpdater = new HeadersUpdater(grid, new UpdateRows(), new UpdateColumns());
                headersUpdater.Update();
            }
        }

        private static DataSet InstanceDataSet(string filePath)
        {
            DataSet dataSet = new DataSet();
            try
            {
                dataSet.ReadXml(filePath);
            }
            catch (XmlException xmlExc)
            {
                DialogResult xmlErrorDialog = MessageBox.Show(xmlExc.Message, "XML File Error",
                    MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error);
                switch (xmlErrorDialog)
                {
                    case DialogResult.Retry:
                    case DialogResult.Ignore:
                        return null;
                    case DialogResult.Abort:
                        Environment.Exit('X');
                        break;
                }
            }

            return dataSet;
        }

        public static void SaveFile(in GridForm grid)
        {
            if (grid.saveFileDialog.ShowDialog().Equals(DialogResult.OK))
            {
                grid.dataGridView.EndEdit();

                string filePath = grid.saveFileDialog.FileName;

                DataTable dataTable = new DataTable("gridData");
                foreach (DataGridViewColumn col in grid.dataGridView.Columns)
                {
                    dataTable.Columns.Add(col.Index.ToString());
                }

                foreach (DataGridViewRow row in grid.dataGridView.Rows)
                {
                    DataRow dtNewRow = dataTable.NewRow();
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        dtNewRow[col.ColumnName] =
                            MathCellsProvider.GetMathCell(row.Cells[int.Parse(col.ColumnName)]).Formula;
                    }
                    dataTable.Rows.Add(dtNewRow);
                }
                dataTable.WriteXml(filePath);
            }
        }
    }
}