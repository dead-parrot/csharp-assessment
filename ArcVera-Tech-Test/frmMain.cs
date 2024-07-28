using Parquet.Schema;
using Parquet;
using System.Data;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using DataColumn = System.Data.DataColumn;
using OxyPlot.Axes;
using System.Text;
using OfficeOpenXml;

namespace ArcVera_Tech_Test
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private async void btnImportEra5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Parquet files (*.parquet)|*.parquet|All files (*.*)|*.*";
                openFileDialog.Title = "Select a Parquet File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    DataTable dataTable = await ReadParquetFileAsync(filePath);
                    dgImportedEra5.DataSource = dataTable;
                    PlotU10DailyValues(dataTable);
                }
            }
        }

        private async Task<DataTable> ReadParquetFileAsync(string filePath)
        {
            using (Stream fileStream = File.OpenRead(filePath))
            {
                using (var parquetReader = await ParquetReader.CreateAsync(fileStream))
                {
                    DataTable dataTable = new DataTable();

                    for (int i = 0; i < parquetReader.RowGroupCount; i++)
                    {
                        using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                        {
                            // Create columns
                            foreach (DataField field in parquetReader.Schema.GetDataFields())
                            {
                                if (!dataTable.Columns.Contains(field.Name))
                                {
                                    Type columnType = field.HasNulls ? typeof(object) : field.ClrType;
                                    dataTable.Columns.Add(field.Name, columnType);
                                }

                                // Read values from Parquet column
                                DataColumn column = dataTable.Columns[field.Name];
                                Array values = (await groupReader.ReadColumnAsync(field)).Data;
                                for (int j = 0; j < values.Length; j++)
                                {
                                    if (dataTable.Rows.Count <= j)
                                    {
                                        dataTable.Rows.Add(dataTable.NewRow());
                                    }
                                    dataTable.Rows[j][field.Name] = values.GetValue(j);
                                }
                            }
                        }
                    }

                    return dataTable;
                }
            }
        }

        private void PlotU10DailyValues(DataTable dataTable)
        {
            var plotModel = new PlotModel { Title = "Daily u10 Values" };
            var lineSeries = new LineSeries { Title = "u10" };

            var groupedData = dataTable.AsEnumerable()
                .GroupBy(row => DateTime.Parse(row["date"].ToString()))
                .Select(g => new
                {
                    Date = g.Key,
                    U10Average = g.Average(row => Convert.ToDouble(row["u10"]))
                })
                .OrderBy(data => data.Date);

            foreach (var data in groupedData)
            {
                lineSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(data.Date), data.U10Average));
            }

            plotModel.Series.Add(lineSeries);
            plotView1.Model = plotModel;
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            DataTable dataTable = (DataTable)dgImportedEra5.DataSource;
            if (dataTable == null)
            {
                MessageBox.Show("No data to export", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog.Title = "Save as CSV";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    ExportDataTableToCsv(dataTable, filePath);
                }
            }
        }

        private void ExportDataTableToCsv(DataTable dataTable, string filePath)
        {
            var start = DateTime.Now;
            var columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            var csv = new StringBuilder();
            csv.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in dataTable.Rows)
            {
                csv.AppendLine(string.Join(",", row.ItemArray));
            }
            File.WriteAllText(filePath, csv.ToString());
            var end = DateTime.Now;
            MessageBox.Show($"Exported {dataTable.Rows.Count} rows to CSV in {(end - start).TotalSeconds} seconds", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            DataTable dataTable = (DataTable)dgImportedEra5.DataSource;
            if (dataTable == null)
            {
                MessageBox.Show("No data to export", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "XLSX files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.Title = "Save as XLSX";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    ExportDataTableToExcel(dataTable, filePath);
                }
            }
        }

        private void ExportDataTableToExcel(DataTable dataTable, string filePath)
        {
            var start = DateTime.Now;
            int rowsPerPage = 1000000;
            int rowsNotImportedYet = dataTable.Rows.Count;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using ExcelPackage package = new ExcelPackage();
            for (int iteration = 0; rowsNotImportedYet > 0; iteration++)
            {
                using DataTable dt = new DataTable();
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    dt.Columns.Add(dataTable.Columns[i].ColumnName, dataTable.Columns[i].DataType);
                }

                int rowsAlreadyImported = dataTable.Rows.Count - rowsNotImportedYet;
                int rowsToImport = Math.Min(rowsPerPage, rowsNotImportedYet);
                for (int i = rowsAlreadyImported; i < rowsAlreadyImported + rowsToImport; i++)
                {
                    dt.ImportRow(dataTable.Rows[i]);
                }
                rowsNotImportedYet -= rowsToImport;

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"Sheet{iteration}");
                var filledRange = worksheet.Cells["A1"].LoadFromDataTable(dt, true);
                int rowCount = dt.Rows.Count;
                var cf = worksheet.ConditionalFormatting.AddExpression(worksheet.Cells[$"A1:E{rowCount+1}"]);
                cf.Formula = "IF($E1<0,1,0)";
                cf.Style.Fill.BackgroundColor.SetColor(Color.Red);
                cf.Style.Font.Color.SetColor(Color.White);

                dt.Clear();
            }
            package.SaveAs(new FileInfo(filePath));

            var end = DateTime.Now;
            MessageBox.Show($"Exported {dataTable.Rows.Count} rows to Excel in {(end - start).TotalSeconds} seconds", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
