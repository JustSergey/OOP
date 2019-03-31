using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace OOP
{
    public class ExcelManager
    {
        public static List<Label> LabelList { get; set; }
        public static List<DataGridView> GridList { get; set; }
        public static SaveFileDialog Dialog { get; set; }

        private static Excel.Application ExcelApp;

        public static void ExportToExcel()
        {
            Dialog.InitialDirectory = "C:";
            Dialog.Title = "Save as Excel File";
            Dialog.FileName = "";
            Dialog.Filter =
                "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx|Excel Files(2013)|*.xlsx";

            if (Dialog.ShowDialog() != DialogResult.Cancel)
            {
                CreateAndFillExcelApp();
            }
        }

        private static void CreateAndFillExcelApp()
        {
            ExcelApp = new Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            FillTables();

            ExcelApp.ActiveWorkbook.SaveCopyAs(Dialog.FileName.ToString());
            ExcelApp.ActiveWorkbook.Saved = true;
            ExcelApp.Quit();
        }

        private static void FillTables()
        {
            ExcelApp.Columns.ColumnWidth = 30;

            var currentRow = 1;

            for (int i = 0; i < LabelList.Count; i++)
            {
                AddLabelToTable(ref currentRow, LabelList[i]);
                AddGridToTable(ref currentRow, GridList[i]);
                currentRow += 2;
            }

            AddGridToTable(ref currentRow, GridList[9]);
            currentRow += 2;
        }

        private static void AddLabelToTable(ref int currentRow,
            Label label, int numberOfWorksheet = 1)
        {
            var ws = (Excel.Worksheet)ExcelApp.Worksheets[numberOfWorksheet];

            var startOfRow = 1;

            ExcelApp.Cells[currentRow, startOfRow] = label.Text.ToString();
            MakeLabelBold(ws, currentRow);
            currentRow++;
        }

        private static void AddGridToTable(ref int currentRow,
            DataGridView dataGridView, int numberOfWorksheet = 1)
        {
            var ws = (Excel.Worksheet)ExcelApp.Worksheets[numberOfWorksheet];

            var rowCount = dataGridView.RowCount;
            var colCount = dataGridView.ColumnCount;

            MakeBordetForTableCells(ws, currentRow, rowCount, colCount);

            //storing headers
            for (int i = 0; i < dataGridView.ColumnCount; i++)
            {
                ExcelApp.Cells[currentRow, i + 1] = dataGridView.Columns[i].HeaderText;
            }
            currentRow++;

            //storing every cell to excel sheet
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                for (int j = 0; j < dataGridView.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + currentRow, j + 1] = dataGridView[j, i].Value;
                }
            }

            currentRow += dataGridView.RowCount;
        }

        private static void MakeLabelBold(Excel.Worksheet ws, int currentRow, int startOfColumn = 1)
        {
            Excel.Range LabelRange = ws.Range[ExcelApp.Cells[currentRow, startOfColumn], ExcelApp.Cells[currentRow, startOfColumn]];
            LabelRange.Font.Bold = true;
        }

        private static void MakeBordetForTableCells(Excel.Worksheet ws, int currentRow, int rowCount, int colCount, int startOfColumn = 1)
        {
            Excel.Range tableRange = ws.Range[ExcelApp.Cells[currentRow, startOfColumn], ExcelApp.Cells[currentRow + rowCount, colCount]];
            tableRange.Borders.Color = System.Drawing.Color.Black.ToArgb();
        }
    }
}
