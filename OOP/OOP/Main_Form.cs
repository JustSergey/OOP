using System;
using System.Net;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Globalization;
//using Application = Microsoft.Office.Interop.Excel.Application;

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

        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            ExcelManager.LabelList = PutTogetherLabels();
            ExcelManager.GridList = PutTogetherDataGridViews();
            ExcelManager.Dialog = SaveToExcelDialog;
            ExcelManager.ExportToExcel();
        }

        private List<Label> PutTogetherLabels()
        {
            var labelList = new List<Label>
            {
                label1,
                label2,
                label3,
                label4,
                label5,
                label6,
                label7,
                label8,
                label9
            };

            return labelList;
        }

        private List<DataGridView> PutTogetherDataGridViews()
        {
            var gridList = new List<DataGridView>
            {
                first_dataGridView,
                second_dataGridView,
                third_dataGridView,
                fourth_dataGridView,
                fifth_dataGridView,
                sixth_dataGridView,
                seventh_dataGridView,
                eighth_dataGridView,
                ninth_dataGridView,
                tenth_dataGridView
            };

            return gridList;
        }
    }
}
