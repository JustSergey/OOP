using System;
using System.Windows.Forms;
using System.IO;
using System.Net;

namespace OOP
{
    public partial class Login_Form : Form
    {
        const string result_file_name = "result.dat";

        public Login_Form()
        {
            InitializeComponent();
            branches_comboBox.Items.AddRange(GetFiliation(DataManager.branches_info_path));
        }

        private void Login_button_Click(object sender, EventArgs e)
        {
            DataManager.BranchIndex = branches_comboBox.SelectedIndex;
            DataManager.QuarterIndex = quarters_comboBox.SelectedIndex;

            if (DataManager.BranchIndex < 0 || DataManager.QuarterIndex < 0)
                return;

            DataManager.BranchName = branches_comboBox.Text;
            Close();
        }

        private void Show_result_button_Click(object sender, EventArgs e)
        {
            if (IsQuarterSelected())
            {
                ConnectAndGetResponse(result_file_name);
                Result_Form result_Form = new Result_Form(result_file_name);
                result_Form.ShowDialog();
                File.Delete(result_file_name);
            }
            else
                ErrorMessage("Для начала выберите квартал");
        }

        private void Save_result_button_Click(object sender, EventArgs e)
        {
            if (IsQuarterSelected())
            {
                saveFileDialog.ShowDialog(this);
                if (saveFileDialog.FileName == "") return;
                ConnectAndGetResponse(saveFileDialog.FileName);
            }
            else
                ErrorMessage("Для начала выберите квартал");
        }

        private bool IsQuarterSelected()
        {
            DataManager.QuarterIndex = quarters_comboBox.SelectedIndex;
            return DataManager.QuarterIndex >= 0;
        }

        private void ErrorMessage(string msg)
        {
            MessageBox.Show(this, msg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static string[] GetFiliation(string Path)
        {
            if (!File.Exists(DataManager.branches_info_path))
                File.WriteAllText(DataManager.branches_info_path, "Филиал");
            return File.ReadAllLines(DataManager.branches_info_path);
        }

        private void ConnectAndGetResponse(string result_path)
        {
            IPEndPoint address = NetManager.GetAddress(DataManager.ip_info_path);
            if (DataManager.ConnectToServer(address))
            {
                DataManager.SendRequest(DataManager.MessageType.GetResult, null);
                DataManager.GetResponse(result_path);
            }
            else
                ErrorMessage("Не удалось подключиться к серверу\nПопробуйте позже");
        }
    }
}
