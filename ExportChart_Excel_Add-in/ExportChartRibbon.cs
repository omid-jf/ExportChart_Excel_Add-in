using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace ExportChart_Excel_Add_in
{
    public partial class ExportChartRibbon
    {
        private void ExportChartRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void aboutBtn_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            aboutBox.ShowDialog();
        }

        private void pdfBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Chart myChart = Globals.ThisAddIn.Application.ActiveChart;

            if (myChart == null)
            {
                MessageBox.Show("Please select a chart!", "Active Chart", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // Getting height and width of the active chart
            Double chHeight = myChart.ChartArea.Height;
            Double chWidth = myChart.ChartArea.Width;

            // Openning a blank Word document
            dynamic app = null;
            dynamic doc = null;
            try
            {
                Type type = Type.GetTypeFromProgID("Word.Application");
                app = Activator.CreateInstance(type);
                //app.Visible = true;

                doc = app.Documents.Add();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }

            // Setting pagesize and margins
            doc.PageSetup.PageHeight = chHeight + 4.7;
            doc.PageSetup.PageWidth = chWidth + 4.7;
            doc.PageSetup.BottomMargin = 0;
            doc.PageSetup.LeftMargin = 0;
            doc.PageSetup.RightMargin = 0;
            doc.PageSetup.TopMargin = 0;

            // Copy chart from Excel
            myChart.ChartArea.Copy();

            // Paste chart to Word, keeping source format
            doc.Range.PasteAndFormat(16);

            // Ask user where to save the PDF
            SaveFileDialog mySaveFileDialog = new SaveFileDialog();
            mySaveFileDialog.Filter = "PDF document (*.pdf)|*.pdf";
            DialogResult myDialogResult = mySaveFileDialog.ShowDialog();
            string filename = null;

            // Save the PDF file
            if (myDialogResult == DialogResult.OK)
            {
                filename = mySaveFileDialog.FileName;
                doc.SaveAs2(filename, 17);
                MessageBox.Show("Exported to " + filename, "Finished Exporting", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                //MessageBox.Show("Please select a location to save the output file.", "Save location", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            doc.Close(0);
            app.Quit();
        }

        private void imgBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Chart myChart = Globals.ThisAddIn.Application.ActiveChart;

            if (myChart == null)
            {
                MessageBox.Show("Please select a chart!", "Active Chart", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Copy chart from Excel
            myChart.ChartArea.Copy();

            // Get Image object from clipboard
            Image image = null;

            if (Clipboard.GetDataObject() != null)
            {
                IDataObject data = Clipboard.GetDataObject();

                if (data.GetDataPresent(DataFormats.Bitmap))
                {
                    image = (Image)data.GetData(DataFormats.Bitmap, true);
                }
            }

            // Ask user where to save the image file and its format
            SaveFileDialog mySaveFileDialog = new SaveFileDialog();
            mySaveFileDialog.Filter = "Bitmap Image (*.bmp)|*.bmp|JPEG Image (*.jpeg)|*.jpeg|Png Image (*.png)|*.png|Tiff Image (*.tiff)|*.tiff";
            DialogResult myDialogResult = mySaveFileDialog.ShowDialog();
            string filename = null;
            string format = null;

            // Save the image file
            if (myDialogResult == DialogResult.OK)
            {
                filename = mySaveFileDialog.FileName;
                format = Path.GetExtension(filename);

                switch (format.ToLower())
                {
                    case ".bmp":
                        image.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);
                        break;

                    case ".jpeg":
                        image.Save(filename, System.Drawing.Imaging.ImageFormat.Jpeg);
                        break;

                    case ".png":
                        image.Save(filename, System.Drawing.Imaging.ImageFormat.Png);
                        break;

                    case ".tiff":
                        image.Save(filename, System.Drawing.Imaging.ImageFormat.Tiff);
                        break;
                }

                MessageBox.Show("Exported to " + filename, "Finished Exporting", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                //MessageBox.Show("Please select a location to save the output file.", "Save location", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
