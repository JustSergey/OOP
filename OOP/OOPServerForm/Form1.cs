﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OOPServerForm
{
    public enum TableType { Extended, ReverseExtended, Ordinary, ReverseOrdinary, NotFilled }
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            DataGridView[] Tables = new DataGridView[] {
                first_dataGridView, second_dataGridView, third_dataGridView,
                fourth_dataGridView, fifth_dataGridView, sixth_dataGridView,
                seventh_dataGridView, eighth_dataGridView, ninth_dataGridView,
                tenth_dataGridView };

            Branch[] branches = new Branch[] {
                new Branch(JustForTests.DeserializeToNewtables("tables0.dat")),
                new Branch(JustForTests.DeserializeToNewtables("tables1.dat")),
            };

            BranchManager branchManager = new BranchManager(branches);
            branchManager.CalcualateBranchesRating();

            InitializeTables(Tables);

            branchManager.FillTables(Tables);
        }

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
            foreach (Control control in Controls)
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
    }
    public class JustForTests
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
                    byte[] bytes = new byte[4];
                    file.Read(bytes, 0, 4);
                    buffer.AddRange(bytes);
                    tables[i].ColumnCount = BitConverter.ToInt32(bytes, 0);
                    file.Read(bytes, 0, 4);
                    buffer.AddRange(bytes);
                    tables[i].RowCount = BitConverter.ToInt32(bytes, 0);
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
    public class BranchManager
    {
        public Branch[] branches;
        private static TableType[] TableTypes = {
            TableType.ReverseExtended,
            TableType.NotFilled,
            TableType.Extended,
            TableType.ReverseExtended,
            TableType.Ordinary,
            TableType.Ordinary,
            TableType.ReverseOrdinary,
            TableType.ReverseExtended,
            TableType.ReverseExtended };// здесь пока не учитывается разделение последней таблицы
        public BranchManager(Branch[] branches)
        {
            this.branches = branches;
        }
        public void CalcualateBranchesRating()
        {
            for (int i = 0; i < TableTypes.Length; i++)
            {
                if ((int)TableTypes[i] == 4)
                    continue;
                else if ((int)TableTypes[i] == 0 || (int)TableTypes[i] == 1)
                    CalculateParametersRaiting(i);
                CalculateFinallRating(i);
            }
        }
        private void CalculateFinallRating(int tableNumber)
        {
            int parameterColumn = branches[0].Tables[tableNumber].ColumnCount - 1;
            var distributionRaiting = GetDistributionRaiting(tableNumber, parameterColumn);
            foreach (var branch in branches)
            {
                var paramValue = GetCellValue(branch.Tables[tableNumber][parameterColumn, 0]);
                branch.Tables[tableNumber].ColumnCount++;// можно просто инкременить count?
                branch.Tables[tableNumber][parameterColumn + 1, 0].Value = distributionRaiting[paramValue];
            }
        }
        private void CalculateParametersRaiting(int tableNumber)// table number от 0
        {
            int parameterColumn = branches[0].Tables[tableNumber].ColumnCount - 1;
            Dictionary<Branch, int> summRatingForBranches = new Dictionary<Branch, int>();

            foreach (var branch in branches)
            {
                summRatingForBranches[branch] = 0;
                branch.Tables[tableNumber].ColumnCount += 2; // место под балл параметра и балл суммарный
            }
            for (var i = 0; i < branches[0].Tables[tableNumber].RowCount; i++) //строки с различными показателями // параметры
            {
                var distributionRaiting = GetDistributionRaiting(tableNumber, parameterColumn, i);
                foreach (var branch in branches)
                {
                    var paramValue = GetCellValue(branch.Tables[tableNumber][parameterColumn, i]);
                    branch.Tables[tableNumber][parameterColumn + 1, i].Value = distributionRaiting[paramValue];
                    summRatingForBranches[branch] += distributionRaiting[paramValue];
                }
            }
            foreach (var branch in branches)
                branch.Tables[tableNumber][parameterColumn + 2, 0].Value = summRatingForBranches[branch];

        }
        private double GetCellValue(DataGridViewCell cell)
        {
            Double.TryParse(cell.Value.ToString(), out double value);
            return value;
        }
        private Dictionary<double, int> GetDistributionRaiting(int tableNumber, int parameterColumn, int parameterRow = 0)// какая
        {
            double[] possibleValues = new double[branches.Length];
            for (var j = 0; j < branches.Length; j++)
                possibleValues[j] = GetCellValue(branches[j].Tables[tableNumber][parameterColumn, parameterRow]);
            Array.Sort(possibleValues);
            if ((int)TableTypes[tableNumber] % 2 == 1)
                Array.Reverse(possibleValues);
            return GetDistributionRaiting(possibleValues);
        }
        private Dictionary<double, int> GetDistributionRaiting(double[] possibleValues)
        {
            Dictionary<double, int> distribution = new Dictionary<double, int>();
            for (var i = 0; i < possibleValues.Length; i++)
            {
                try
                {
                    distribution.Add(possibleValues[i], distribution.Count + 1);
                }
                catch { }
            }
            return distribution;
        }
        public void FillTables(DataGridView[] Tables)// тест 
        {
            for (int tableNum = 0; tableNum < Tables.Length - 1; tableNum++)// БЕЗ последней таблицы и итоговой 
                for (int branchNum = 0; branchNum < branches.Length; branchNum++)
                    for (var i = 1; i < Tables[tableNum].ColumnCount; i++)// колонки
                        for (var j = 0; j < branches[branchNum].Tables[tableNum].RowCount; j++)//строки
                        {
                            if (((int)TableTypes[tableNum] == 0 || (int)TableTypes[tableNum] == 1) && i == 1)
                                continue;
                            Tables[tableNum][i, j + branchNum * branches[branchNum].Tables[tableNum].RowCount].Value = branches[branchNum].Tables[tableNum][i - 1, j].Value;
                        }
            FillFinallTable(Tables[Tables.Length - 1]);
        }
        private void FillFinallTable(DataGridView table)
        {
            double[] ratingsSums = new double[branches.Length];
            for (var i = 0; i < branches.Length; i++)
            {
                double summOfRatingScores = 0;
                for (var j = 0; j < branches[i].Tables.Length; j++)
                {
                    var columnCount = branches[i].Tables[j].ColumnCount;
                    table[3 + i, j].Value = branches[i].Tables[j][columnCount - 1, 0].Value;
                    summOfRatingScores += GetCellValue(table[3 + i, j]) * GetCellValue(table[2, j]);
                }
                table[3 + i, branches[i].Tables.Length].Value = summOfRatingScores;
                ratingsSums[i] = summOfRatingScores;
            }
            var disributionRating = GetDistributionRaiting(ratingsSums);
            for (var i = 0; i < branches.Length; i++)
                table[i+3, branches[i].Tables.Length + 1].Value = disributionRating[GetCellValue(table[i+3, branches[i].Tables.Length])];
        }
    }
    public struct Branch
    {
        public readonly DataGridView[] Tables;
        public Branch(DataGridView[] tables)
        {
            Tables = tables;
        }
    }
}

