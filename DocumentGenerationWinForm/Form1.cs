using Microsoft.Office.Interop.Word;
using System.Collections.Specialized;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Point = System.Drawing.Point;
using Font = System.Drawing.Font;


namespace DocumentGenerationWinForm
{
    public partial class Form1 : Form
    {
        string runtimeFileName;
        Word.Application WordApp;
        Excel.Application ExcelApp;
        string selectedTypeOfProject;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CustomResponsiveLayout();
            CreateFolderInPublicFolder("DocumentGenerationWinForm");
            DatasourceCheck();
            FillComboBox();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            CustomResponsiveLayout();
        }

        private void Form1_ResizeEnd(object sender, EventArgs e)
        {
            CustomResponsiveLayout();
        }
        private void Form1_ControlAdded(object sender, ControlEventArgs e)
        {
            CustomResponsiveLayout();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            AcceptClick();
        }
        private void CustomResponsiveLayout()
        {
            var comboBox1XLocation = this.Location.X + this.Width * (.25);
            var comboBox1YLocation = this.Location.Y + this.Height * (.33);
            comboBox1.Location = this.PointToClient(new Point(Convert.ToInt32(comboBox1XLocation),
                Convert.ToInt32(comboBox1YLocation)));
            comboBox1.Height = Convert.ToInt32(this.Height * (.05));
            comboBox1.Width = Convert.ToInt32(this.Width * (.50));
            comboBox1.Font = SetFontSize(Convert.ToInt32(this.Width * (.01)));

            var textBox1XLocation = this.Location.X + this.Width * (.25);
            var textBox1YLocation = this.Location.Y + this.Height * (.33) + comboBox1.Height + 20;
            textBox1.Location = this.PointToClient(new Point(Convert.ToInt32(textBox1XLocation),
                Convert.ToInt32(textBox1YLocation)));
            textBox1.Height = Convert.ToInt32(comboBox1.Height);
            textBox1.Width = Convert.ToInt32(this.Width * (.40));
            textBox1.Font = comboBox1.Font;

            var button1XLocation = this.Location.X + this.Width * (.65);
            var button1YLocation = this.Location.Y + this.Height * (.33) + comboBox1.Height + 20;
            button1.Location = this.PointToClient(new Point(Convert.ToInt32(button1XLocation),
                Convert.ToInt32(button1YLocation)));
            button1.Height = Convert.ToInt32(comboBox1.Height);
            button1.Width = Convert.ToInt32(this.Width * (.10));
            button1.Font = comboBox1.Font;

            var button2XLocation = this.Location.X + this.Width * (.25);
            var button2YLocation = this.Location.Y + this.Height * (.33) + comboBox1.Height + textBox1.Height + 80;
            button2.Location = this.PointToClient(new Point(Convert.ToInt32(button2XLocation),
                Convert.ToInt32(button2YLocation)));
            button2.Height = Convert.ToInt32(comboBox1.Height);
            button2.Width = Convert.ToInt32(this.Width * (.10));
            button2.Font = comboBox1.Font;

            var button3XLocation = this.Location.X + this.Width * (.35);
            var button3YLocation = this.Location.Y + this.Height * (.33) + comboBox1.Height + textBox1.Height + 80;
            button3.Location = this.PointToClient(new Point(Convert.ToInt32(button3XLocation),
                Convert.ToInt32(button3YLocation)));
            button3.Height = Convert.ToInt32(comboBox1.Height);
            button3.Width = Convert.ToInt32(this.Width * (.10));
            button3.Font = comboBox1.Font;
        }

        private void SearchClick(object sender, EventArgs e)
        {
            DialogResult dialogResult = openFileDialog1.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                button2.Enabled = true;
            }
        }

        private void AcceptClick()
        {
            //Create destination folder named with current datetime
            string publicFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            string folderFilePath = Path.Combine(publicFolderPath, "DocumentGenerationWinForm");
            string currentDateTime = DateTime.Now.ToString("dddd, dd MMMM yyyy HH.mm.ss");
            string destinationFolderPath = System.IO.Path.Combine(folderFilePath, currentDateTime);
            if (!ExistsFolder(destinationFolderPath))
            {
                CreateFolder(destinationFolderPath);
            }
            //Assign names of selected documents to collection
            StringCollection documentsForSelectedTypeOfProjectCollection = GetNamesOfSelectedDocuments();
            //Assign information provided in the Excel to dictionary
            Dictionary<string, string> dataDictionary = GetTwoFirstColumnsOfExcelSheet(openFileDialog1.FileName,"data");

            ProcessWordsOfSelectedTypeOfProject(documentsForSelectedTypeOfProjectCollection, 
                dataDictionary, folderFilePath, destinationFolderPath,
                "Los archivos han sido generados con éxito");

            //UI-effects
            textBox1.Text = String.Empty;
            button1.Enabled = false;
            button2.Enabled = false;

        }

        private void ProcessWordsOfSelectedTypeOfProject(StringCollection documentsForSelectedTypeOfProjectCollection,
            Dictionary<string, string> dataDictionary, string folderFilePath, 
            string destinationFolderFilePath, string successMessage)
        {
            object missing = System.Reflection.Missing.Value;
            Word.Application WordApp = new Word.Application();

            foreach (string selectedWord in documentsForSelectedTypeOfProjectCollection) {
                string selectedWordFilePath = System.IO.Path.Combine(folderFilePath, selectedWord + ".docx");
                //Open app and document
                Word.Document wordDoc = WordApp.Documents.Open(selectedWordFilePath, ReadOnly: true, Visible: false);

                //Include operations here
                //Copy content
                wordDoc.ActiveWindow.Selection.WholeStory();
                wordDoc.ActiveWindow.Selection.Copy();

                //Paste content in new document, modify fields and save it
                string newDocumentFilePath = System.IO.Path.Combine(destinationFolderFilePath, selectedWord + " - " + selectedTypeOfProject + ".docx");
                var newDocument = new Microsoft.Office.Interop.Word.Document();
                newDocument.ActiveWindow.Selection.Paste();

                //Iterate over items in dataDictionary changing ocurrences of information
                Microsoft.Office.Interop.Word.Range wordRange = newDocument.Content;
                foreach (KeyValuePair<string, string> item in dataDictionary)
                {
                    wordRange.Find.Execute(FindText: item.Key.ToString(), ReplaceWith: item.Value.ToString(), Replace: Word.WdReplace.wdReplaceAll);
                }
                newDocument.SaveAs(newDocumentFilePath);
                Object newDocumentFilePathObject = (object)newDocumentFilePath;

                string newDocumentFilePathPDF = System.IO.Path.Combine(destinationFolderFilePath, selectedWord + " - " + selectedTypeOfProject + ".pdf");
                Object newDocumentFilePathPDFObject = (object)newDocumentFilePathPDF;
                newDocument.SaveAs(ref newDocumentFilePathPDFObject, WdSaveFormat.wdFormatPDF, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

                //Close document and app
                wordDoc.Close();
                NAR(wordDoc);
                newDocument.Close();
                NAR(newDocument);
            }
            WordApp.Quit();
            NAR(WordApp);

            //Notify success
            MessageBox.Show(successMessage);
        }
        private Dictionary<string, string> GetTwoFirstColumnsOfExcelSheet(string filePath, string sheetName)
        {
            //Open app and book
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook excelBook = ExcelApp.Workbooks.Open(filePath);

                //Open sheet
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelBook.Worksheets[sheetName];

            //Include operations here
            Dictionary<string, string> dataDictionary = dataDictionary = new Dictionary<string, string>();
            int maxRow = excelWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int row = 1;
            while (row <= maxRow)
            {
                Excel.Range mCell = excelWorksheet.Cells[row, 1];
                Excel.Range nCell = excelWorksheet.Cells[row, 2];
                string mCellValue = (string)mCell.Value2;
                string nCellValue = nCell.Text.ToString();
                if (nCellValue is double)
                {
                    MessageBox.Show(nCellValue);
                    nCellValue = Convert.ToString(double.Parse(nCell.Value2));
                    MessageBox.Show(nCellValue);
                }

                if (mCellValue == null || mCellValue == String.Empty || nCellValue == null || nCellValue == String.Empty)
                {
                    row++;
                    continue;
                }
                else
                {
                    dataDictionary.Add(mCellValue, nCellValue);
                    row++;
                    continue;
                }
            }

            //Close book and app
            excelBook.Close();
                NAR(excelBook);
                ExcelApp.Quit();
                NAR(ExcelApp);

            return dataDictionary;
        }
        private StringCollection GetNamesOfSelectedDocuments()
        {
            //Initialize collection
            StringCollection documentsForSelectedTypeOfProjectCollection = new StringCollection();
            string folderName = "DocumentGenerationWinForm";
            string fileName = "DocumentGenerationWinForm.xlsx";
            string newFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), folderName,
                fileName);
            //Open app and book
            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook excelBook = ExcelApp.Workbooks.Open(newFilePath);

            //Open sheet
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelBook.Worksheets["data"];

            //Include operations here
            int maxRow = excelWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int maxCol = excelWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            int selectedTypeOfProjectId = comboBox1.SelectedIndex + 1;

            int col = 2;
            while (col <= maxCol)
            {
                Excel.Range mCell = excelWorksheet.Cells[selectedTypeOfProjectId + 1, col];
                string testValue = (string)mCell.Value2;
                if (testValue == null || testValue == String.Empty || testValue == "no")
                {
                    col++;
                    continue;
                }
                else
                {
                    Excel.Range dCell = excelWorksheet.Cells[1, col];
                    string documentToBeIncluded = (string)dCell.Value2;
                    documentsForSelectedTypeOfProjectCollection.Add(documentToBeIncluded);
                    col++;
                    continue;
                }
            }

            //Close book and app
            excelBook.Close();
            NAR(excelBook);
            ExcelApp.Quit();
            NAR(ExcelApp);
            
            return documentsForSelectedTypeOfProjectCollection;
        }
        private void SetSelectedTypeOfProject(object sender, EventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            selectedTypeOfProject = (string)cmb.SelectedItem;
            button1.Enabled = true;
        }
        private void FillComboBox()
        {
            string folderName = "DocumentGenerationWinForm";
            string fileName = "DocumentGenerationWinForm.xlsx";
            string newFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), folderName,
                fileName);
            //Open app and book
            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook excelBook = ExcelApp.Workbooks.Open(newFilePath);

            //Open sheet
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelBook.Worksheets["data"];

            //Include operations here
            int maxRow = excelWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int maxCol = excelWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            StringCollection typeOfProjectsStringCollection = new StringCollection();

            // Elements of first row to a collection
            int row = 1;
            while (row <= maxRow)
            {
                Excel.Range mCell = excelWorksheet.Cells[row, 1];
                if (mCell.Value2 == null || mCell.Value2 == String.Empty)
                {
                    row++;
                    continue;
                }
                else
                {
                    string typeOfProjectToBeAdded = (string)mCell.Value2;
                    typeOfProjectsStringCollection.Add(typeOfProjectToBeAdded);
                    row++;
                    continue;
                }
            }

            foreach (string s in typeOfProjectsStringCollection)
            {
                comboBox1.Items.Add(s);
            }


            //Close book and app
            excelBook.Close();
            NAR(excelBook);
            ExcelApp.Quit();
            NAR(ExcelApp);
        }
        private void CreateFolderInPublicFolder(string folderName)
        {
            string newFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), folderName);
            if (!ExistsFolder(newFolderPath))
            {
                CreateFolder(newFolderPath);
            }
        }


        private void DatasourceCheck()
        {
            string folderName = "DocumentGenerationWinForm";
            string fileName = "DocumentGenerationWinForm.xlsx";
            string newFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), folderName,
                fileName);
            if (!ExistsFile(newFilePath))
            {
                MessageBox.Show("Datasource not found. This application cannot run without it.");
                this.Close();
            }
        }
        internal static void NAR(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null;
            }
        }

        #region REPOSITORY
        /*REPOSITORY*/
        private Boolean ExistsFile(string filePath)
        {
            if (File.Exists(filePath)) return true;
            else return false;
        }

        /*Check if folder exists*/
        private Boolean ExistsFolder(string folderPath)
        {
            if (Directory.Exists(folderPath)) return true;
            else return false;
        }
        /*Create folder*/
        private void CreateFolder(string destinationFolderPath)
        {
            if (!Directory.Exists(destinationFolderPath))
                Directory.CreateDirectory(destinationFolderPath);
        }
        private Font SetFontSize(int fontSize)
        {
            Font myFont;
            try
            {
                myFont = new Font(FontFamily.GenericSansSerif, fontSize);
            }
            catch (Exception ex)
            {
                myFont = new Font(FontFamily.GenericSansSerif, fontSize);
            }
            return myFont;
        }

        #endregion

    }
}