using QRCoder;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.IO;


namespace QR_Master
{
    public partial class Interface : Form
    {
        // Generated Serial Number (Single QR)
        private string serial;

        public Interface()
        {
            InitializeComponent();
            LoadItemsToComboBox();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            serial = Utility.generateserialnumbers();
            string qrText = $"{serial},{cmbCategoryS.Text},{cmbTypeS.Text},{cmbItemS.Text},{cmbSizeS.Text},{cmbDepartmentS.Text}";

            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrText, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(5);

            pbQRArea.SizeMode = PictureBoxSizeMode.CenterImage;
            pbQRArea.Image = qrCodeImage;
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
            if (pbQRArea.Image != null)
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "PNG Image|*.png|JPEG Image|*.jpg|Bitmap Image|*.bmp";
                    saveFileDialog.Title = "Save QR Code Image";
                    saveFileDialog.FileName = $"QRCode-{serial}-{cmbCategoryS.Text}.png";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Save the image in the chosen format
                        switch (saveFileDialog.FilterIndex)
                        {
                            case 1:
                                pbQRArea.Image.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Png);
                                break;
                            case 2:
                                pbQRArea.Image.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                break;
                            case 3:
                                pbQRArea.Image.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Bmp);
                                break;
                        }
                    }
                }

                // Save to Log
                Utility.LogData(@"C:\QRMateData\Logs\Generate_Log_s.txt", serial, cmbCategoryS.Text, cmbTypeS.Text, cmbItemS.Text, cmbSizeS.Text, cmbDepartmentS.Text, "1", "1");

                MessageBox.Show("QR code downloaded successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Please generate a QR code first.", "No QR Code", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            cmbCategoryS.SelectedIndex = -1;
            cmbTypeS.SelectedIndex = -1;
            cmbItemS.SelectedIndex = -1;
            cmbSizeS.SelectedIndex = -1;
            cmbDepartmentS.SelectedIndex = -1;
            pbQRArea.Image = null;
        }

        private void btnGeneratePDF_Click(object sender, EventArgs e)
        {
            int quantity;
            int startingSequence;

            if (int.TryParse(nudQuantity.Text, out quantity) && quantity > 0 &&
                int.TryParse(txtSequence.Text, out startingSequence))
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "PDF Document|*.pdf";
                    saveFileDialog.Title = "Save QR Codes as PDF";
                    saveFileDialog.FileName = $"QRCodes-{DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")}-{cmbCategoryC.Text}.pdf";


                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Initiate Progressbar
                        Utility.UpdateProgressBar(progressBar, 0, quantity);

                        PdfDocument pdf = new PdfDocument();
                        PdfPage page = pdf.AddPage();
                        XGraphics gfx = XGraphics.FromPdfPage(page);

                        int qrPerRow = 5; // Number of QR codes per row
                        int qrPerColumn = 5; // Number of QR codes per column
                        int qrSize = 100; // Size of each QR code
                        int margin = 20; // Margin between QR codes

                        int x = margin;
                        int y = margin;
                        int count = 0;

                        for (int i = 0; i < quantity; i++)
                        {
                            int sequenceNumber = startingSequence + i;
                            string serialNumber = Utility.generateserialnumbers();
                            string qrText = $"{sequenceNumber},{serialNumber},{cmbCategoryC.Text},{cmbTypeC.Text},{cmbItemC.Text},{cmbSizeC.Text},{cmbDepartmentC.Text}";
                            QRCodeGenerator qrGenerator = new QRCodeGenerator();
                            QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrText, QRCodeGenerator.ECCLevel.Q);
                            QRCode qrCode = new QRCode(qrCodeData);
                            Bitmap qrCodeImage = qrCode.GetGraphic(5);

                            using (MemoryStream stream = new MemoryStream())
                            {
                                qrCodeImage.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                                stream.Position = 0;
                                XImage xImage = XImage.FromStream(stream);
                                gfx.DrawImage(xImage, x, y, qrSize, qrSize); // Adjust for size and position as needed
                            }

                            string text = $"{cmbCategoryC.Text}-{cmbTypeC.Text}-{cmbItemC.Text} {sequenceNumber}";
                            XFont font = new XFont("Arial", 8);
                            XSize textSize = gfx.MeasureString(text, font);
                            double textX = x + (qrSize - textSize.Width) / 2;

                            gfx.DrawString(text, font, XBrushes.Black, new XPoint(textX, y + qrSize + 10));

                            x += qrSize + margin;
                            count++;

                            if (count % qrPerRow == 0)
                            {
                                x = margin;
                                y += qrSize + margin + 20; // Adjust for text height
                            }

                            if (count % (qrPerRow * qrPerColumn) == 0 && i < quantity - 1)
                            {
                                page = pdf.AddPage();
                                gfx = XGraphics.FromPdfPage(page);
                                x = margin;
                                y = margin;
                            }

                            // Update progress bar
                            Utility.UpdateProgressBar(progressBar, i + 1, quantity);
                        }

                        pdf.Save(saveFileDialog.FileName);

                        // Save to Log
                        Utility.LogData(@"C:\QRMateData\Logs\Generate_Log_c.txt", "-", cmbCategoryC.Text, cmbTypeC.Text, cmbItemC.Text, cmbSizeC.Text, cmbDepartmentC.Text, nudQuantity.Text, txtSequence.Text);

                        progressBar.Visible = false;

                        MessageBox.Show("PDF saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please enter valid quantity and starting sequence.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            ClearInputValuesC();
        }

        void ClearInputValuesC()
        {
            cmbCategoryC.SelectedIndex = -1;
            cmbTypeC.SelectedIndex = -1;
            cmbItemC.SelectedIndex = -1;
            cmbSizeC.SelectedIndex = -1;
            cmbDepartmentC.SelectedIndex = -1;
            nudQuantity.Text = string.Empty;
            txtSequence.Text = string.Empty;
        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            ClearInputValuesC();
        }

        private void LoadItemsToComboBox()
        {
            var products = Utility.ReadExcelData(@"C:\QRMateData\data_source.xlsx");

            // Single
            Utility.PopulateComboBoxes(cmbCategoryS, cmbTypeS, cmbItemS, cmbSizeS, products);
            Utility.LoadItemsToComboBoxes(cmbDepartmentS, @"C:\QRMateData\Department.txt");

            // Collection
            Utility.PopulateComboBoxes(cmbCategoryC, cmbTypeC, cmbItemC, cmbSizeC, products);
            Utility.LoadItemsToComboBoxes(cmbDepartmentC, @"C:\QRMateData\Department.txt");
        }

        private void btnLoadS_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\QRMateData\Logs\Generate_Log_s.txt";
            Utility.LoadDataToDataGridView(filePath, dgvLoadS);
        }

        private void btnLoadC_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\QRMateData\Logs\Generate_Log_c.txt";
            Utility.LoadDataToDataGridView(filePath, dgvLoadC);
        }

        private void btnExportS_Click(object sender, EventArgs e)
        {
            Utility.ExportData(dgvLoadS, progressBar);
        }

        private void btnExportC_Click(object sender, EventArgs e)
        {
            Utility.ExportData(dgvLoadC, progressBar);
        }

        private void Interface_Load(object sender, EventArgs e)
        {
            progressBar.Visible = false;
        }
    }
}
