using System;
using System.Windows.Forms;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;

namespace SearchInExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog diag = new FolderBrowserDialog();
            if (diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string folder = diag.SelectedPath;  //selected folder path

                textBox1.Text = folder;

            }
        }
        public WorkbookPart ImportExcel(string filepath)
        {
            try
            {

                using (FileStream fs = File.Open(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    MemoryStream m_ms = new MemoryStream();
                    fs.CopyTo(m_ms);

                    SpreadsheetDocument m_Doc = SpreadsheetDocument.Open(m_ms, false);

                    return m_Doc.WorkbookPart;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.Message + ex.StackTrace);
            }
            return null;
        }

        public void GetIndexBySearch(string search, string filepath)
        {
            logBox.Text = logBox.Text + Environment.NewLine + filepath;
            string[] letters = { "A", 
                "B", 
                "C", 
                "D",
                "E",
                "F",
                "G",
                "H",
                "I",
                "J",
                "K",
                "L",
                "M",
                "N",
                "O",
                "P",
                "Q",
                "R",
                "S",
                "T",
                "U",
                "V",
                "W",
                "X",
                "Y",
                "Z"};
            WorkbookPart workbookPart = ImportExcel(filepath);
            var sheets = workbookPart.Workbook.Descendants<Sheet>().ToList();

            foreach(var sheet in sheets) {

                string index = "Not Found";

                if (sheet != null)
                {

                    logBox.Text = logBox.Text + Environment.NewLine + "SheetName: "+sheet.Name;
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    var rows = worksheetPart.Worksheet.Descendants<Row>().ToList();

                    // Remove the header row
                    if(rows.Count > 0)
                    {
                        //rows.RemoveAt(0);
                    }

                    //logBox.Text = logBox.Text + Environment.NewLine + rows.Count;


                    foreach (var row in rows)
                    {
                        var cellss = row.Elements<Cell>().ToList();

                        //logBox.Text = logBox.Text + Environment.NewLine + cellss.Count;

                        foreach (var cell in cellss)
                        {

                            if (cell.StyleIndex != null)
                            {
                                var value = cell.InnerText;

                                //logBox.Text = logBox.Text + Environment.NewLine + cell.InlineString;

                                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                int value_int = 0;

                                //hay un problema en este parse
                                if (!int.TryParse(value, out value_int)) value_int = 0;

                                
                                if (value_int != 0)
                                {
                                    try {
                                        value = stringTable.SharedStringTable.ElementAt(value_int).InnerText;
                                    } catch (ArgumentOutOfRangeException e){
                                        value = value_int.ToString();
                                    }

                                } else { 
                                    value = ""; 
                                }
                                bool isFound = value.Trim().ToLower().Contains(search.Trim().ToLower());

                                //logBox.Text = logBox.Text + Environment.NewLine + value;

                                if (isFound)
                                {
                                    int letterindx = (int)(GetColumnIndex(cell.CellReference) - 1);
                                    index = $"[{letters[letterindx]}{row.RowIndex}]";
                                    logBox.Text = logBox.Text + Environment.NewLine + cell.CellReference.InnerText + " " + value;
                                }
                            }

                        }
                    }
                }

                if (index.Equals("Not Found"))
                {
                    logBox.Text = logBox.Text + Environment.NewLine + index;
                }

            } 

            
        }

        private static int? GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return null;
            }

            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier = mulitplier * 26;
            }

            return columnNumber + 1;
        }

        private void searchButton_Click(object sender, EventArgs e)
        {

            logBox.Text = "";

            DirectoryInfo d = new DirectoryInfo(textBox1.Text); //Assuming Test is your Folder

            FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting Text files

            foreach (FileInfo file in Files)
            {
                
                GetIndexBySearch(searchBox.Text, file.FullName);
            }
        }
    }
}
