using System;
using System.Net;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace OOP
{
    public partial class Main_Form : Form
    {
        private TableManager tableManager;
        private DataGridView[] Tables;

        private string TableFileName
        {
            get => "t" + DataManager.BranchIndex + "q" + DataManager.QuarterIndex + ".dat";
        }

        public Main_Form()
        {
            InitializeComponent();

            Tables = GetADataGridView();
            TableManager.InitializeTables(Tables);
            tableManager = new TableManager(Tables);
        }

        private DataGridView[] GetADataGridView()
        {
            return new DataGridView[] {
                first_dataGridView, second_dataGridView, third_dataGridView,
                fourth_dataGridView, fifth_dataGridView, sixth_dataGridView,
                seventh_dataGridView, eighth_dataGridView, ninth_dataGridView,
                tenth_dataGridView };
        }

        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            tableManager.FillTable((DataGridView)sender);
        }

        private void Save_send_button_Click(object sender, EventArgs e)
        {
            DataManager.Serialize(Tables, TableFileName);

            IPEndPoint address = NetManager.GetAddress(DataManager.ip_info_path);
            if (!DataManager.ConnectToServer(address))
                MessageError("Не удалось подключиться к серверу\r\nПопробуйте позже");
            else
                DataManager.SendRequest(DataManager.MessageType.SendFile, TableFileName);
        }

        private void MessageError(string msg)
        {
            MessageBox.Show(this, msg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            Login_Form login_Form = new Login_Form();
            login_Form.ShowDialog(this);
            if (DataManager.BranchIndex < 0 || DataManager.QuarterIndex < 0)
                Close();

            Text += ": Данные (" + DataManager.BranchName + ") за " + (DataManager.QuarterIndex + 1) + " квартал";
            DataManager.Deserialize(Tables, TableFileName);
        }

        private void Main_Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (DataManager.BranchIndex >= 0 && DataManager.QuarterIndex >= 0)
                DataManager.Serialize(Tables, TableFileName);
        }

        private static void MakeBordetForTableCells(Excel.Application ExcelApp, Excel.Worksheet ws, 
                                    int currentRow, int rowCount, int colCount, int startOfColumn = 1)
        {
            Excel.Range tableRange = ws.Range[ExcelApp.Cells[currentRow, startOfColumn], ExcelApp.Cells[currentRow + rowCount, colCount]];
            tableRange.Borders.Color = System.Drawing.Color.Black.ToArgb();
        }

        private void AddLabelToTable(Excel.Application ExcelApp, ref int currentRow,
            Label label, int numberOfWorksheet = 1)
        {
            var ws = (Excel.Worksheet)ExcelApp.Worksheets[numberOfWorksheet];

            ExcelApp.Cells[currentRow, 1] = label.Text.ToString();
            MakeLabelBold(ExcelApp, ws, currentRow);
            currentRow++;
        }

        private static void MakeLabelBold(Excel.Application ExcelApp, Excel.Worksheet ws, int currentRow, int startOfColumn = 1)
        {
            Excel.Range LabelRange = ws.Range[ExcelApp.Cells[currentRow, startOfColumn], ExcelApp.Cells[currentRow, startOfColumn]];
            LabelRange.Font.Bold = true;
        }

        private void AddGridToTable(Excel.Application ExcelApp, ref int currentRow,
            DataGridView dataGridView, int numberOfWorksheet = 1)
        {
            Excel.Worksheet ws = (Excel.Worksheet)ExcelApp.Worksheets[numberOfWorksheet];

            MakeBordetForTableCells(ExcelApp, ws, currentRow, dataGridView.RowCount, dataGridView.ColumnCount);

            storingHeaders(ExcelApp, currentRow, dataGridView);
            storingEveryCell(ExcelApp, currentRow, dataGridView);

            currentRow += dataGridView.RowCount;
        }

        private static void storingEveryCell(Excel.Application ExcelApp, int currentRow, DataGridView dataGridView)
        {
            for (int i = 0; i < dataGridView.RowCount; i++)
                for (int j = 0; j < dataGridView.ColumnCount; j++)
                    ExcelApp.Cells[i + currentRow, j + 1] = dataGridView[j, i].Value;
        }

        private static void storingHeaders(Excel.Application ExcelApp, int currentRow, DataGridView dataGridView)
        {
            for (int i = 0; i < dataGridView.ColumnCount; i++)
                ExcelApp.Cells[currentRow, i + 1] = dataGridView.Columns[i].HeaderText;
            currentRow++;
        }

        private List<Label> PutTogetherLabels()
        {
            List<Label> labelList = new List<Label>
            {
                label1, label2, label3,
                label4, label5, label6,
                label7, label8, label9
            };
            return labelList;
        }

        private List<DataGridView> PutTogetherDataGridViews()
        {
            List<DataGridView> gridList = new List<DataGridView>
            {
                first_dataGridView, second_dataGridView, third_dataGridView,
                fourth_dataGridView, fifth_dataGridView, sixth_dataGridView,
                seventh_dataGridView, eighth_dataGridView, ninth_dataGridView,
                tenth_dataGridView
            };
            return gridList;
        }

        private void FillTables(Excel.Application ExcelApp)
        {
            ExcelApp.Columns.ColumnWidth = 30;

            int currentRow = 1;

            for (int i = 0; i < PutTogetherLabels().Count; i++)
            {
                AddLabelToTable(ExcelApp, ref currentRow, PutTogetherLabels()[i]);
                AddGridToTable(ExcelApp, ref currentRow, PutTogetherDataGridViews()[i]);
                currentRow += 2;
            }

            AddGridToTable(ExcelApp, ref currentRow, PutTogetherDataGridViews()[9]);
            currentRow += 2;
        }

        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            SaveToExcelDialog.InitialDirectory = "C:";
            SaveToExcelDialog.Title = "Save as Excel File";
            SaveToExcelDialog.FileName = "";
            SaveToExcelDialog.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx|Excel Files(2013)|*.xlsx";

            if (SaveToExcelDialog.ShowDialog() != DialogResult.Cancel)
            {
                var ExcelApp = new Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing);

                FillTables(ExcelApp);

                ExcelApp.ActiveWorkbook.SaveCopyAs(SaveToExcelDialog.FileName.ToString());
                ExcelApp.ActiveWorkbook.Saved = true;
                ExcelApp.Quit();
            }
        }
    }
}
