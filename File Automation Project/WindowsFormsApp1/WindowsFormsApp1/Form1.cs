using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private string FilePath;
        private string Path;
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c: \\";
            openFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
               FilePath = openFileDialog1.FileName;
               label3.Text = "Excel file is selected";
            }           
        }
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.InitialDirectory = "c: \\";
            openFileDialog2.Filter = "Word Files (.Docx)|*.Docx|All Files (*.*)|*.*";
            openFileDialog2.FilterIndex = 1;
            openFileDialog2.RestoreDirectory = true;
            if(openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                Path = openFileDialog2.FileName;
                label4.Text = "Word File is selected";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            string excelpath = FilePath;
            string wordpath = Path;

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelpath, ReadOnly: false, Editable: true);
            Excel.Worksheet worksheet = workbook.Worksheets[1];

            Word.Application wordApp = new Word.Application();
            Word.Document document = wordApp.Documents.Add();
            try
            {
                System.Threading.Thread.Sleep(600);

                worksheet.UsedRange.Copy();

                document.Application.Selection.Range.PasteSpecial();
                System.Threading.Thread.Sleep(600);
                document.SaveAs2(wordpath);
                MessageBox.Show("File is Automated");
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");
            }
            finally
            {
                // Clean up
                workbook.Save();
                workbook.Close(SaveChanges: true);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(worksheet);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
            }
        }
    }
}
