using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;

namespace QR_Master
{
    internal class Utility
    {
        public static string generateserialnumbers()
        {
            return Guid.NewGuid().ToString().Substring(0, 10);
        }

        public static void LogData(string filePath, string serial, string category, string type, string item, string size, string department, string quantity, string sequence)
        {
            string logFilePath = filePath;
            string logEntry = $"{DateTime.Now}, {serial}, {category}, {type}, {item}, {size}, {department}, {quantity}, {sequence}";

            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while logging data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void LoadDataToDataGridView(string filePath, DataGridView dgv)
        {
            if (File.Exists(filePath))
            {
                var lines = File.ReadAllLines(filePath);
                var data = lines.Select(line => line.Split(','))
                                .Where(parts => parts.Length >= 9) // Ensure there are at least 9 parts
                                .Select(parts => new
                                {
                                    DateTime = parts[0].Trim(),
                                    SerialNumber = parts[1].Trim(),
                                    Category = parts[2].Trim(),
                                    Type = parts[3].Trim(),
                                    Item = parts[4].Trim(),
                                    Size = parts[5].Trim(),
                                    Division = parts[6].Trim(),
                                    Quantity = parts[7].Trim(),
                                    SequenceStart = parts[8].Trim()
                                }).ToList();

                dgv.DataSource = data;
            }
            else
            {
                MessageBox.Show("File not found.");
            }
        }

        public static void ExportData(DataGridView dgv, ProgressBar progressBar)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv";

            string date = DateTime.Now.ToString("dd-MM-yyyy");
            sfd.FileName = $"QR_Mate-Export-{date}.xlsx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Initiate progressbar
                UpdateProgressBar(progressBar, 0, dgv.Rows.Count);

                try
                {
                    if (System.IO.Path.GetExtension(sfd.FileName).ToLower() == ".csv")
                    {
                        // Export to CSV
                        StringBuilder sb = new StringBuilder();

                        string[] columnNames = dgv.Columns.Cast<DataGridViewColumn>().
                                                      Select(column => "\"" + column.HeaderText + "\"").
                                                      ToArray();
                        sb.AppendLine(string.Join(",", columnNames));

                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            DataGridViewRow row = dgv.Rows[i];
                            string[] cells = row.Cells.Cast<DataGridViewCell>().
                                                  Select(cell => "\"" + cell.Value + "\"").
                                                  ToArray();
                            sb.AppendLine(string.Join(",", cells));
                        }

                        System.IO.File.WriteAllText(sfd.FileName, sb.ToString());
                    }
                    else
                    {
                        // Export to Excel
                        var excelApp = new Excel.Application();
                        excelApp.Workbooks.Add();

                        Excel._Worksheet workSheet = excelApp.ActiveSheet;

                        for (int i = 0; i < dgv.Columns.Count; i++)
                        {
                            workSheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
                        }

                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            DataGridViewRow row = dgv.Rows[i];
                            for (int j = 0; j < row.Cells.Count; j++)
                            {
                                workSheet.Cells[i + 2, j + 1] = row.Cells[j].Value;

                                // Update progress bar
                                UpdateProgressBar(progressBar, i + 1, dgv.Rows.Count);
                            }
                        }

                        workSheet.SaveAs(sfd.FileName);
                        excelApp.Quit();
                    }

                    progressBar.Visible = false;

                    MessageBox.Show("File saved successfully.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static void UpdateProgressBar(ProgressBar progressBar, int value, int maxValue)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => UpdateProgressBar(progressBar, value, maxValue)));
            }
            else
            {
                progressBar.Minimum = 0;
                progressBar.Maximum = maxValue;
                progressBar.Value = value;
                progressBar.Visible = true;
            }
        }

        public static void LoadItemsToComboBoxes(ComboBox comboBox, string filePath)
        {
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                comboBox.Items.AddRange(lines);
            }
            else
            {
                MessageBox.Show("The file does not exist.", "No File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Populate Comboboxes
        public static List<Product> ReadExcelData(string filePath)
        {
            var products = new List<Product>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    products.Add(new Product
                    {
                        Category = worksheet.Cells[row, 1].Text,
                        Type = worksheet.Cells[row, 2].Text,
                        Item = worksheet.Cells[row, 3].Text,
                        Size = worksheet.Cells[row, 4].Text
                    });
                }
            }

            return products;
        }

        public static void PopulateComboBoxes(ComboBox categoryComboBox, ComboBox typeComboBox, ComboBox itemComboBox, ComboBox sizeComboBox, List<Product> products)
        {
            var categories = products.Select(p => p.Category).Distinct().ToList();
            categoryComboBox.DataSource = categories;
            categoryComboBox.SelectedIndex = -1; 

            categoryComboBox.SelectedIndexChanged += (s, e) =>
            {
                var selectedCategory = categoryComboBox.SelectedItem?.ToString();
                if (selectedCategory != null)
                {
                    var types = products.Where(p => p.Category == selectedCategory).Select(p => p.Type).Distinct().ToList();
                    typeComboBox.DataSource = types;
                    typeComboBox.SelectedIndex = -1; 
                }
            };

            typeComboBox.SelectedIndexChanged += (s, e) =>
            {
                var selectedCategory = categoryComboBox.SelectedItem?.ToString();
                var selectedType = typeComboBox.SelectedItem?.ToString();
                if (selectedCategory != null && selectedType != null)
                {
                    var items = products.Where(p => p.Category == selectedCategory && p.Type == selectedType).Select(p => p.Item).Distinct().ToList();
                    itemComboBox.DataSource = items;
                    itemComboBox.SelectedIndex = -1; 
                }
            };

            itemComboBox.SelectedIndexChanged += (s, e) =>
            {
                var selectedCategory = categoryComboBox.SelectedItem?.ToString();
                var selectedType = typeComboBox.SelectedItem?.ToString();
                var selectedItem = itemComboBox.SelectedItem?.ToString();
                if (selectedCategory != null && selectedType != null && selectedItem != null)
                {
                    var sizes = products.Where(p => p.Category == selectedCategory && p.Type == selectedType && p.Item == selectedItem).Select(p => p.Size).Distinct().ToList();
                    sizeComboBox.DataSource = sizes;
                    sizeComboBox.SelectedIndex = -1; 
                }
            };
        }
    }
}
