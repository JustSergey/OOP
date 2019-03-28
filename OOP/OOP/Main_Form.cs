using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using System.Text;
using System.Threading;

namespace OOP
{
    public partial class Main_Form : Form
    {
        public const string ip_info_path = "ip.ini";
        public const string branches_info_path = "branches.inf";

        private TableManager tableManager;
        private DataGridView[] Tables;

        private string TableFileName {
            get => "t" + DataManager.BranchIndex + "q" + DataManager.QuarterIndex + ".dat"; }
        
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
    }
}
