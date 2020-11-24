using System;
using System.Windows.Forms;

namespace MyExcel
{
    internal partial class GridForm : Form
    {
        public GridForm()
        {
            InitializeComponent();

            // Enabling double bufferisation for DataGridView
            typeof(DataGridView).InvokeMember(
                "DoubleBuffered",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.SetProperty,
                null,
                dataGridView,
                new object[] { true });

            MathCellsProvider.GetInstance.SetGridForm(this);
            InitialGridSetup();
        }

        public void ClearGrid()
        {
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
        }

        public void ClearOutputBoxes()
        {
            AddressBox.Text = string.Empty;
            FormulaBox.Text = string.Empty;
        }

        private void InitialGridSetup()
        {
            MathCellsProvider.GetInstance.ResetMathCells();
            ClearGrid();
            ClearOutputBoxes();

            const int initN = 10;
            const int initM = 10;

            dataGridView.RowCount = initN;
            dataGridView.ColumnCount = initM;

            HeadersUpdater headersUpdater = new HeadersUpdater(this, new UpdateRows(), new UpdateColumns());
            headersUpdater.Update();

            OnGridMathCellsProvider.InitEmptyCells(this);
        }

        private void NewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InitialGridSetup();
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FilesHandler.OpenFile(this);
        }

        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FilesHandler.SaveFile(this);
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void HelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            const string helpMess = "This program can be used as a lightweight Microsoft Excel.\n" +
                                    "It allows you to write and overwrite data in cells and execute " +
                                    "different types of calculations, written in infix format. \n" +
                                    "You can use references to another cells with =RiCi notation." +
                                    "For example: =R1C1 + 2 + R3C5\n" +
                                    "Mathematical functions powered by mXParser. You can find more information " +
                                    "about supported functions on official site: http://mathparser.org/ \n" +
                                    "Try to input such expression as: =(51 - 1) * R2C1 / 2";
            MessageBox.Show(helpMess, "Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            const string aboutMess = "Faculty of Computer Science and Cybernetics\n" +
                                     "Group: K-24\n" +
                                     "Student: Danylchenko Alexander\n" +
                                     "Year: 2019";
            MessageBox.Show(aboutMess, "About", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void ButtonAddRow_Click(object sender, EventArgs e)
        {
            GridModifier.AddRow(this);
        }

        private void ButtonAddColumn_Click(object sender, EventArgs e)
        {
            GridModifier.AddColumn(this);
        }

        private void ButtonInsertRow_Click(object sender, EventArgs e)
        {
            GridModifier.InsertRow(this);
            ClearOutputBoxes();
        }

        private void ButtonInsertColumn_Click(object sender, EventArgs e)
        {
            GridModifier.InsertColumn(this);
            ClearOutputBoxes();
        }

        private void DeleteRowButton_Click(object sender, EventArgs e)
        {
            GridModifier.DeleteRow(this);
            ClearOutputBoxes();
        }

        private void DeleteColumnButton_Click(object sender, EventArgs e)
        {
            GridModifier.DeleteColumn(this);
            ClearOutputBoxes();
        }

        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            MathCell mCell = MathCellsProvider.GetMathCell(dataGridView.CurrentCell);
            AddressBox.Text = mCell.OwnAddress;
            FormulaBox.Text = mCell.Formula;
        }

        private void DataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView.BeginEdit(true);
        }

        private void DataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            MathCell mCell = MathCellsProvider.GetMathCell(dataGridView.CurrentCell);
            dataGridView.CurrentCell.Value = mCell.Formula;
        }

        private bool IsGridChanged { get; set; }
        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell inputCell = dataGridView.CurrentCell;
            MathCell inputMathCell = MathCellsProvider.GetInstance.ProcessCellInput(inputCell);

            dataGridView.CurrentCell.Value = inputMathCell.Value;
            FormulaBox.Text = inputMathCell.Formula;

            if (MathCellsProvider.GetInstance.GetUsedCells().Contains(inputMathCell))
            {
                MathCellsProvider.GetInstance.RemoveCell(inputMathCell);
            }
            MathCellsProvider.GetInstance.AddCell(inputMathCell);
            IsGridChanged = true;
        }

        private void GridForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsGridChanged)
            {
                DialogResult closeSaveResult = MessageBox.Show("Do you want to save last changes?",
                    "Save File", MessageBoxButtons.YesNoCancel);

                if (closeSaveResult == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
                else if (closeSaveResult == DialogResult.Yes)
                {
                    FilesHandler.SaveFile(this);
                }
            }
        }
    }
}