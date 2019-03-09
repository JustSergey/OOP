using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using System.Text;

namespace OOP
{
    public partial class Main_Form : Form
    {
        const string tables_data_path = "tables.dat";
        const string info_ini_path = "info.ini";
        const string _ip = "127.0.0.1";
        const int _port = 3123;
        private TableManager tableManager;
        private IPEndPoint address;
        private UdpClient client = new UdpClient();
        private DataGridView[] Tables;

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
            TableManager.Serialize(Tables, tables_data_path);
            if (!SendOnServer(tables_data_path, 14))
            {
                MessageBox.Show(this, "Не удалось отправить данные на сервер\n\rПопробуйте позже",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            TableManager.Deserialize(Tables, tables_data_path);
            ParseIp(info_ini_path);
        }

        private void ParseIp(string file_path)
        {
            //Parse file for ip and port
            if (!File.Exists(file_path))
            {
                File.WriteAllText(file_path, _ip + ":" + _port.ToString());
            }
            string[] ip_port = File.ReadAllText(file_path).Split(':');

            IPAddress ip;
            int port;
            if (IPAddress.TryParse(ip_port[0], out ip) && int.TryParse(ip_port[1], out port))
                address = new IPEndPoint(ip, port);
        }

        private bool SendOnServer(string file_path, int filiation_id)
        {
            byte[] file = File.ReadAllBytes(file_path);
            byte[] id = BitConverter.GetBytes(filiation_id);
            client.Send(id, id.Length, address);
            int bytes_sent = client.Send(file, file.Length, address);
            return bytes_sent == file.Length;
        }
    }

    public class TableManager
    {
        private delegate void TableFiller();
        private readonly Dictionary<DataGridView, TableFiller> tablesFillers = new Dictionary<DataGridView, TableFiller>();

        public TableManager(DataGridView[] tables)
        {
            TableFiller[] TableFillers = GetTableFillers(tables);
            for (int i = 0; i < tables.Length; i++)
                tablesFillers[tables[i]] = TableFillers[i];
        }

        public static void InitializeTables(DataGridView[] Tables)
        {
            Tables[0].Rows.Add("Промышленность", "0", "0", "0");
            Tables[0].Rows.Add("Население", "0", "0", "0");
            Tables[0].Rows.Add("Бюджет", "0", "0", "0");
            Tables[0].Rows.Add("ОПП, ЖКХ и др.", "0", "0", "0");
            Tables[0].Rows.Add("Прочее", "0", "0", "0");
            Tables[1].Rows.Add("0", "0", "0", "0");
            Tables[2].Rows.Add("Физическое", "0", "0", "0");
            Tables[2].Rows.Add("Юридическое", "0", "0", "0");
            Tables[3].Rows.Add("Физическое", "0", "0", "1", "0");
            Tables[3].Rows.Add("Юридическое", "0", "0", "1", "0");
            Tables[4].Rows.Add("0", "0", "0", "0", "0");
            Tables[5].Rows.Add("0");
            Tables[6].Rows.Add("0", "0", "0");
            Tables[7].Rows.Add("Инциденты", "0", "0", "0");
            Tables[7].Rows.Add("Техника безопасности", "0", "0", "0");
            Tables[8].Rows.Add("0", "0", "0");
            Tables[9].Rows.Add("0", "0", "0");
        }

        public static void Serialize(DataGridView[] Tables, string file_path)
        {
            using (FileStream file = File.Create(file_path))
            {
                for (int i = 0; i < Tables.Length; i++)
                {
                    for (int y = 0; y < Tables[i].RowCount; y++)
                    {
                        for (int x = 0; x < Tables[i].ColumnCount; x++)
                        {
                            byte[] buffer = Encoding.Unicode.GetBytes(Tables[i][x, y].Value.ToString());
                            file.Write(buffer, 0, buffer.Length);
                            file.WriteByte(0x02);
                            file.WriteByte(0xA8);
                        }
                    }
                }
            }
        }

        public static void Deserialize(DataGridView[] Tables, string file_path)
        {
            if (!File.Exists(file_path))
                return;
            using (FileStream file = File.OpenRead(file_path))
            {
                for (int i = 0; i < Tables.Length; i++)
                {
                    for (int y = 0; y < Tables[i].RowCount; y++)
                    {
                        for (int x = 0; x < Tables[i].ColumnCount; x++)
                        {
                            List<byte> buffer = new List<byte>(byte.MaxValue);
                            byte[] bytes = new byte[2];
                            while (file.Read(bytes, 0, 2) == 2)
                            {
                                if (bytes[0] == 0x02 && bytes[1] == 0xA8)
                                    break;
                                buffer.AddRange(bytes);
                            }
                            Tables[i][x, y].Value = Encoding.Unicode.GetString(buffer.ToArray());
                        }
                    }
                }
            }
        }

        public void FillTable(DataGridView table)
        {
            tablesFillers[table]();
        }

        private bool TryParseCells(DataGridViewCell[] cells, out int[] tableParams)
        {
            tableParams = new int[cells.Length];            
            for (int i = 0; i < cells.Length; i++)
            {
                if (!int.TryParse(cells[i].Value.ToString(), out tableParams[i]) || tableParams[i] < 0) // нет проверки равенста с нгулем!(есть случаи, где 0 допустимо)
                    return false;
            }
            return true;
        }

        private bool TryParseCells(DataGridViewCell[] cells, out double[] tableParams)
        {
            tableParams = new double[cells.Length];
            for (int i = 0; i < cells.Length; i++)
            {
                if (!double.TryParse(cells[i].Value.ToString(), out tableParams[i]) || tableParams[i] < 0)// нет проверки равенста с нгулем!(есть случаи, где 0 допустимо)
                    return false;
            }
            return true;
        }        
        private TableFiller[] GetTableFillers(DataGridView[] tables)
        {           
            TableFiller[] tablesFillers = new TableFiller[tables.Length];
            tablesFillers[0] = () =>
            {
                DataGridViewCell[] cells;
                for (int i = 0; i < tables[0].RowCount; i++)
                {
                    cells = new DataGridViewCell[] { tables[0][1, i], tables[0][2, i] };
                    if (!TryParseCells(cells, out double[] parameters) || parameters[1] <= 0)
                        continue;
                    tables[0][3, i].Value = parameters[0] * 365 / parameters[1];
                }
            };

            tablesFillers[1] = () =>
            {
                DataGridViewCell[] cells = { tables[1][0, 0], tables[1][1, 0] };
                if (!TryParseCells(cells, out double[] parameters) || parameters[1] <= 0)
                    return;
                tables[1][2, 0].Value = parameters[0] / parameters[1] * 100;
                var cellValue = (double)tables[1][2, 0].Value;
                if (cellValue >= 120)
                    tables[1][3, 0].Value = 6;
                else if (cellValue >= 110)
                    tables[1][3, 0].Value = 5;
                else if (cellValue >= 100)
                    tables[1][3, 0].Value = 4;
                else if (cellValue >= 95)
                    tables[1][3, 0].Value = 3;
                else if (cellValue >= 90)
                    tables[1][3, 0].Value = 2;
                else
                    tables[1][3, 0].Value = 1;
            };

            tablesFillers[2] = () =>
            {
                int documentsCount = 0;
                DataGridViewCell[] cells;
                for (var i = 0; i < tables[2].RowCount; i++)
                {
                    cells = new DataGridViewCell[] { tables[2][1, i], tables[2][2, i] };
                    if (!TryParseCells(cells, out int[] parameters) || parameters[1] <= 0)
                        continue;
                    tables[2][3, i].Value = (double)parameters[0] / parameters[1];
                    documentsCount += parameters[1];
                }
                tables[6][1, 0].Value = documentsCount;
                FillTable(tables[6]);
            };

            tablesFillers[3] = () =>
            {
                DataGridViewCell[] cells;
                for (var i = 0; i < tables[3].RowCount; i++)
                {
                    cells = new DataGridViewCell[] { tables[3][1, i], tables[3][2, i], tables[3][3, i] };
                    if (!TryParseCells(cells, out int[] parameters) || parameters[1] <= 0)
                        continue;
                    tables[3][4, i].Value = (double)parameters[0] / parameters[1] * parameters[2] * 100;
                }
                tables[4][0, 0].Value = tables[3][2, 0].Value;
                tables[4][1, 0].Value = tables[3][2, 1].Value;
                FillTable(tables[4]);
            };

            tablesFillers[4] = () =>
            {
                DataGridViewCell[] cells = { tables[4][0, 0], tables[4][1, 0], tables[4][3, 0] };
                if (!TryParseCells(cells, out int[] parameters))
                    return;
                if (parameters[2] > 0)
                    tables[4][4, 0].Value = (double)(parameters[0] + parameters[1]) / parameters[2];
            };

            tablesFillers[5] = () => { };

            tablesFillers[6] = () =>
            {
                DataGridViewCell[] cells = { tables[6][0, 0], tables[6][1, 0] };
                if (!TryParseCells(cells, out int[] parameters) || parameters[1] <= 0)
                    return;
                tables[6][2, 0].Value = (double)parameters[0] / parameters[1];
            };

            tablesFillers[7] = () =>
            {
                DataGridViewCell[] cells;
                for (int i = 0; i < tables[7].Rows.Count; i++)
                {
                    cells = new DataGridViewCell[] { tables[7][1, i], tables[7][2, i] };
                    if (!TryParseCells(cells, out int[] parameters) || parameters[1] <= 0)
                        continue;
                    tables[7][3, i].Value = (double)parameters[0] / parameters[1];
                }
            };

            tablesFillers[8] = () =>
            {
                DataGridViewCell[] cells = { tables[8][0, 0], tables[8][1, 0] };
                if (!TryParseCells(cells, out int[] parameters) || parameters[1] <= 0)
                    return;
                tables[8][2, 0].Value = (double)parameters[0] / parameters[1] * 100;
            };

            tablesFillers[9] = () =>
            {
                DataGridViewCell[] cells = { tables[9][0, 0], tables[9][1, 0] };
                if (!TryParseCells(cells, out int[] parameters) || parameters[1] <= 0)
                    return;
                tables[9][2, 0].Value = (double)parameters[0] / parameters[1] * 100;
            };

            return tablesFillers;
        }
    }
}
