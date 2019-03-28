using System;
using System.Net;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace OOP
{
    public partial class Main_Form : Form
    {
        public const string ip_info_path = "ip.ini";
        public const string branches_info_path = "branches.inf";

        private TableManager tableManager;
        private DataGridView[] Tables;

        private string TableFileName
        {
            get => "t" + DataManager.BranchIndex + "q" + DataManager.QuarterIndex + ".dat";
        }

        public Main_Form()
        {
            InitializeComponent();

            Tables = new DataGridView[] {
                first_dataGridView, second_dataGridView, third_dataGridView,
                fourth_dataGridView, fifth_dataGridView, sixth_dataGridView,
                seventh_dataGridView, eighth_dataGridView, ninth_dataGridView,
                tenth_dataGridView };

            TableManager.InitializeTables(Tables);
            tableManager = new TableManager(Tables);
        }

        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            tableManager.FillTable((DataGridView)sender);
        }

        private void Save_send_button_Click(object sender, EventArgs e)
        {
            DataManager.Serialize(Tables, TableFileName);

            IPEndPoint address = DataManager.ParseIp(ip_info_path);
            if (!DataManager.ConnectToServer(address))
            {
                MessageBox.Show(this, "Не удалось подключиться к серверу\r\nПопробуйте позже",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DataManager.SendRequest(DataManager.MessageType.SendFile, TableFileName);
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

        private void first_dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void SaveToExcelDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {


        }

        private void fifth_dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            SaveToExcelDialog.InitialDirectory = "C:";
            SaveToExcelDialog.Title = "Save as Excel File";
            SaveToExcelDialog.FileName = "";
            SaveToExcelDialog.Filter =
                "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx|Excel Files(2013)|*.xlsx";

            if (SaveToExcelDialog.ShowDialog() != DialogResult.Cancel)
            {
                var ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing);

                ExcelApp.Columns.ColumnWidth = 20;

                //storing headers
                for (int i = 1; i < first_dataGridView.ColumnCount + 1; i++)
                {
                    ExcelApp.Cells[1, i] = first_dataGridView.Columns[i - 1].HeaderText;
                }

                //storing every cell to excel sheet
                for (int i = 0; i < first_dataGridView.RowCount; i++)
                {
                    for (int j = 0; j < first_dataGridView.ColumnCount; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = first_dataGridView[j, i].Value.ToString();
                    }
                }

                ExcelApp.ActiveWorkbook.SaveCopyAs(SaveToExcelDialog.FileName.ToString());
                ExcelApp.ActiveWorkbook.Saved = true;
                ExcelApp.Quit();
            }
        }
    }
}
