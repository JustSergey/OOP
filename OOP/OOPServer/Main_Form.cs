using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OOPServer
{
    public enum TableType { Extended, ReverseExtended, Ordinary, ReverseOrdinary, NotFilled }
    public partial class Main_Form : Form
    {
        public Main_Form(string[] branch_file_names)
        {
            InitializeComponent();

            DataGridView[] Tables = new DataGridView[] {
                first_dataGridView, second_dataGridView, third_dataGridView,
                fourth_dataGridView, fifth_dataGridView, sixth_dataGridView,
                seventh_dataGridView, eighth_dataGridView, ninth_dataGridView,
                tenth_dataGridView };

            Branch[] branches = new Branch[branch_file_names.Length];
            for (int i = 0; i < branches.Length; i++)
                branches[i] = new Branch(DataManager.DeserializeToNewtables(branch_file_names[i]));

            BranchManager branchManager = new BranchManager(branches);
            branchManager.CalcualateBranchesRating();
            InitializeTables(Tables);
            branchManager.FillTables(Tables);
        }

        //public Main_Form(string[][] branch_file_names)
        //{
        //    InitializeComponent();

        //    DataGridView[] Tables = new DataGridView[] {
        //        first_dataGridView, second_dataGridView, third_dataGridView,
        //        fourth_dataGridView, fifth_dataGridView, sixth_dataGridView,
        //        seventh_dataGridView, eighth_dataGridView, ninth_dataGridView,
        //        tenth_dataGridView };

        //    Branch[][] branches = new Branch[4][];

        //    for (int i = 0; i < branches.Length; i++)
        //    {
        //        branches[i] = new Branch[branches[i].Length];//(DataManager.DeserializeToNewtables(branch_file_names[i]));
        //        for (int j = 0; j < branches[i].Length; j++)
        //            branches[i][j] = new Branch(DataManager.DeserializeToNewtables(branch_file_names[i][j]));
        //    }



        //    BranchManager branchManager = new BranchManager(branches);
        //    branchManager.CalcualateBranchesRating();
        //    InitializeTables(Tables);
        //    branchManager.FillTables(Tables);
        //}
        public void InitializeTables(DataGridView[] Tables)
        {
            for (int i = 1; i < 7; i++)
            {
                Tables[0].Rows.Add(i, "Промышленность", "0", "0", "0", "0", "0", "0");
                Tables[0].Rows.Add(i, "Население", "0", "0", "0", "0", "0", "0");
                Tables[0].Rows.Add(i, "Бюджет", "0", "0", "0", "0", "0", "0");
                Tables[0].Rows.Add(i, "ОПП, ЖКХ и др.", "0", "0", "0", "0", "0", "0");
                Tables[0].Rows.Add(i, "Прочее", "0", "0", "0", "0", "0", "0");
                Tables[1].Rows.Add(i, "0", "0", "0", "0");
                Tables[2].Rows.Add(i, "Физическое", "0", "0", "0", "0", "0", "0");
                Tables[2].Rows.Add(i, "Юридическое", "0", "0", "0", "0", "0", "0");
                Tables[3].Rows.Add(i, "Физическое", "0", "0", "1", "0", "0", "0", "0");
                Tables[3].Rows.Add(i, "Юридическое", "0", "0", "1", "0", "0", "0", "0");
                Tables[4].Rows.Add(i, "0", "0", "0", "0", "0", "0");
                Tables[5].Rows.Add(i, "0", "0");
                Tables[6].Rows.Add(i, "0", "0", "0", "0");
                Tables[7].Rows.Add(i, "Инциденты", "0", "0", "0", "0", "0", "0");
                Tables[7].Rows.Add(i, "Техника безопасности", "0", "0", "0", "0", "0", "0");
                Tables[8].Rows.Add(i, "Коффициент текучести кадров", "0", "0", "0", "0", "0", "0");
                Tables[8].Rows.Add(i, "Качество обучения", "0", "0", "0", "0", "0", "0");
            }

            string[] parametrs_text = new string[9];
            foreach (Control control in panel.Controls)
            {
                if (control is Label)
                {
                    string name = control.Name;
                    if (name.Length != 6)
                        continue;
                    if (!int.TryParse(name.Substring(5), out int index))
                        continue;
                    parametrs_text[index - 1] = control.Text.Remove(control.Text.Length - 1);
                }
            }
            for (int i = 1; i < parametrs_text.Length + 1; i++)
                Tables[9].Rows.Add(i, parametrs_text[i - 1], "0", "0", "0", "0", "0", "0", "0");
            double[] parameterCoefficients = { 1.5, 1.5, 1, 1, 1, 0.5, 0.5, 0.5, 0.5 };
            for (int i = 0; i < parameterCoefficients.Length; i++)
                Tables[9][2, i].Value = parameterCoefficients[i];
            Tables[9].Rows.Add("", "Сумма баллов с учетом веса", "", "0", "0", "0", "0", "0", "0");
            Tables[9].Rows.Add("", "Итоговое местов рейтинге", "", "0", "0", "0", "0", "0", "0");
        }

        private void DataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            DataGridView dataGrid = (DataGridView)sender;
            if ((e.ColumnIndex == 0 ||
                e.ColumnIndex == dataGrid.ColumnCount - 1 ||
                e.ColumnIndex == dataGrid.ColumnCount - 2) &&
                (e.RowIndex < dataGrid.RowCount - 1 && e.RowIndex >= 0))
            {
                if (dataGrid[0, e.RowIndex].Value.ToString() == dataGrid[0, e.RowIndex + 1].Value.ToString())
                    e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                else
                    e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.Single;
            }
        }

        private void DataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            Common_dataGridView_CellFormatting(sender, e);
            DataGridView dataGrid = (DataGridView)sender;
            if ((e.ColumnIndex == 0 ||
                   e.ColumnIndex == dataGrid.ColumnCount - 1 ||
                   e.ColumnIndex == dataGrid.ColumnCount - 2) &&
                   e.RowIndex > 0)
            {
                if (dataGrid[0, e.RowIndex].Value.ToString() == dataGrid[0, e.RowIndex - 1].Value.ToString())
                    e.CellStyle.ForeColor = e.CellStyle.BackColor;
            }
        }

        private void Common_dataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridView dataGrid = (DataGridView)sender;
            if (e.RowIndex >= 0)
            {
                if (int.TryParse(dataGrid[0, e.RowIndex].Value.ToString(), out int value))
                {
                    if (value % 2 == 0)
                        e.CellStyle.BackColor = Color.Lavender;
                }
                else
                    e.CellStyle.BackColor = Color.MistyRose;
            }
        }

        private void Main_Form_Shown(object sender, EventArgs e)
        {
            Bitmap bmp = new Bitmap(panel.Width, panel.Height);
            Rectangle bounds = new Rectangle(0, 0, bmp.Width, bmp.Height);
            panel.DrawToBitmap(bmp, bounds);
            bmp.Save("result.dat", System.Drawing.Imaging.ImageFormat.Png);
            Close();
        }
    }
    public class DataManager 
    {
        public static DataGridView[] DeserializeToNewtables(string file_path)
        {
            DataGridView[] tables = new DataGridView[10];
            if (!File.Exists(file_path))
                return tables;
            using (FileStream file = File.OpenRead(file_path))
            {
                for (int i = 0; i < tables.Length; i++)
                {
                    tables[i] = new DataGridView();
                    List<byte> buffer = new List<byte>(byte.MaxValue);
                    byte[] bytes = new byte[8];
                    file.Read(bytes, 0, 8);
                    buffer.AddRange(bytes);
                    tables[i].ColumnCount = BitConverter.ToInt32(bytes, 0);
                    tables[i].RowCount = BitConverter.ToInt32(bytes, 4);
                    for (int y = 0; y < tables[i].RowCount; y++)
                    {
                        for (int x = 0; x < tables[i].ColumnCount; x++)
                        {
                            buffer = new List<byte>(byte.MaxValue);
                            bytes = new byte[2];
                            while (file.Read(bytes, 0, 2) == 2)
                            {
                                if (bytes[0] == 0x02 && bytes[1] == 0xA8)
                                    break;
                                buffer.AddRange(bytes);
                            }
                            tables[i][x, y].Value = Encoding.Unicode.GetString(buffer.ToArray());
                        }
                    }
                }
            }
            return StandardizeTables(tables);
        }

        private static DataGridView[] StandardizeTables(DataGridView[] tables)
        {
            tables[8] = MergeTables(tables[8], tables[9]);
            Array.Resize<DataGridView>(ref tables, tables.Length - 1);
            tables[8].Columns.Insert(0, new DataGridViewColumn(tables[8][0, 0]));
            return tables;
        }
        private static DataGridView MergeTables(DataGridView table1, DataGridView table2)
        {
            var table3 = new DataGridView();
            table3.ColumnCount = table1.ColumnCount;
            table3.RowCount = table1.RowCount + table2.RowCount;
            for (var i = 0; i < table1.ColumnCount; i++)
                for (var j = 0; j < table1.RowCount; j++)
                    table3[i, j].Value = table1[i, j].Value;
            var rowCount = table1.RowCount;
            for (var i = 0; i < table2.ColumnCount; i++)
                for (var j = rowCount; j < rowCount + table2.RowCount; j++)
                    table3[i, j].Value = table2[i, j - rowCount].Value;
            return table3;
        }
    }
    
}

