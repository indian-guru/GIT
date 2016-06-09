using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ExcelImageImporter
{
    public partial class Importer : Form
    {
        public Importer()
        {
            InitializeComponent();
        }

        private void buttonExtract_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBoxFilePath.Text))
            {
                return;
            }
            if (string.IsNullOrEmpty(textBoxImagePath.Text))
            {
                return;
            }

            Extract();
        }

        private void Extract()
        {
            Excel.Application application = new Excel.Application {Visible = false};
            Excel.Workbook workbook = application.Workbooks.Open(textBoxFilePath.Text);
            Excel.Worksheet worksheet = workbook.Sheets[1];

            Excel.Pictures pics = worksheet.Pictures(Missing.Value) as Excel.Pictures;
            if (pics != null)
            {
                progressBar.Maximum = pics.Count;
                for (var i = 1; i <= pics.Count; i++)
                {
                    progressBar.Value = i;

                    try
                    {
                        pics.Item(i).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                        var image = Clipboard.GetImage();
                        string imageName = worksheet.Cells[pics.Item(i).TopLeftCell.Row, 3].Formula;
                        image?.Save(Path.Combine(textBoxImagePath.Text, imageName + ".bmp"));
                    }
                    catch (Exception)
                    {
                        // ignored
                    }
                }
            }

            workbook.Close();
            application.Quit();
            MessageBox.Show(@"Image Extraction completed.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            Process.Start(textBoxImagePath.Text);
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog document = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                Filter = @"Excel files (*.xls;*.xlsx)|*.xls; *.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if (document.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (document.CheckFileExists)
                    {
                        textBoxFilePath.Text = document.FileName.Trim();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void buttonImagePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog
            {
                RootFolder = Environment.SpecialFolder.Desktop,
                ShowNewFolderButton = true
            };

            if (folder.ShowDialog() == DialogResult.OK)
            {
                textBoxImagePath.Text = folder.SelectedPath.Trim();
            }


        }
    }
}
