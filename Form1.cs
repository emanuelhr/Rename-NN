using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using System.IO;
using System;

namespace WinFormsApp1

{
    public partial class Form1 : Form
    {

        static string ChangeFileExtension(string filePath, string newExtension)
        {
            // Check if the file path is not empty and has an extension
            if (!string.IsNullOrEmpty(filePath) && Path.HasExtension(filePath))
            {
                // Create a new file path with the specified extension
                string directory = Path.GetDirectoryName(filePath);
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                string newFilePath = Path.Combine(directory, fileNameWithoutExtension + "." + newExtension);
                return newFilePath;
            }

            // Return the original file path if it doesn't have an extension or is empty
            return filePath;
        }


        public bool ReplaceNN(string oldNN, string newNN, string path)
        {
            string input = "ID,NN,'" + oldNN + "'";
            string replaced = "ID,NN,'" + newNN + "'";
            string extractedText = null;

            //Project

            string searchPrefix = "SACHM=";
            string[] lines = File.ReadAllLines(path);

            // Search for the line starting with the specified prefix
            foreach (string line in lines)
            {
                if (line.StartsWith(searchPrefix))
                {
                    // Extract the text after the prefix
                    extractedText = line.Substring(searchPrefix.Length);

                    // Display the result
                    Console.WriteLine("Extracted Text: " + extractedText);

                }
            }

            //Table

            string tablePath = CurrentFolder(path) + "\\" + extractedText;

            string text = File.ReadAllText(tablePath);
            bool exists = text.Contains(oldNN);
            string newText = text.Replace(input, replaced);
            if (exists)
            {
                File.WriteAllText(tablePath, newText);
            }
            Console.WriteLine("Project done " + path);
            return exists;

        }

        //Pronalazi text u single quotama
        static string SelectTextWithinSingleQuotes(string input)
        {
            // Use a regular expression to match text within single quotes
            System.Text.RegularExpressions.Match match = Regex.Match(input, @"'([^']*)'");

            // Check if a match is found
            if (match.Success)
            {
                // Extract and return the text within single quotes
                return match.Groups[1].Value;
            }

            // Return an empty string if no match is found
            return string.Empty;
        }


        public string CurrentFolder(string path)
        {
            // Example file path
            string filePath = path;

            // Get the current folder from the file path
            string currentFolder = Path.GetDirectoryName(filePath);

            return currentFolder;
        }


        public Form1()
        {
            InitializeComponent();
        }

        private void browseInput_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "D:\\23dlibs"; // You can set the initial directory here
                openFileDialog.Filter = "Xls (*.xls)|*.xls|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the selected file's path
                    string selectedFilePath = openFileDialog.FileName;

                    // Now you can do something with the selected file path, for example, display it in a TextBox
                    textBox1.Text = selectedFilePath;
                }
            }
        }

        private void run_Click(object sender, EventArgs e)
        {
            string filePath = textBox1.Text;
            List<string> currentPath = new List<string>();
            List<string> oldNN = new List<string>();
            List<string> newNN = new List<string>();
            List<string> replaced = new List<string>();
            string newPath = CurrentFolder(textBox1.Text) + "\\replaceCheck.xls";


            if (File.Exists(filePath))
            {
                try
                {
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        HSSFWorkbook workbook = new HSSFWorkbook(fileStream);
                        ISheet sheet = workbook.GetSheetAt(0); // Assuming you are working with the first sheet

                        // Iterate through rows
                        for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                        {
                            IRow row = sheet.GetRow(rowIndex);

                            if (row != null)
                            {
                                // Iterate through cells
                                for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                                {
                                    ICell cell = row.GetCell(cellIndex);

                                    if (cell != null)
                                    {

                                        switch (cellIndex)
                                        {
                                            case 0:
                                                currentPath.Add(cell.ToString());

                                                break;
                                            case 1:
                                                oldNN.Add(cell.ToString());
                                                break;
                                            case 2:
                                                newNN.Add(cell.ToString());
                                                break;
                                            default:
                                                Console.WriteLine("Error, it should not have more than 3 columns");
                                                break;
                                        }





                                    }
                                }
                                Console.WriteLine();
                            }
                        }



                        for (int i = 0; i < currentPath.Count; i++)
                        {

                            if (ReplaceNN(oldNN[i], newNN[i], currentPath[i]))
                            {
                                replaced.Add("Replaced");
                            }
                            else
                            {
                                replaced.Add("Not Replaced");
                            }

                        }



                    }

                    // Create a new workbook
                    HSSFWorkbook workbook2 = new HSSFWorkbook();

                    // Create a worksheet named "Sheet1"
                    ISheet sheet2 = workbook2.CreateSheet("Check");

                    int rowCount = currentPath.Count;

                    IRow startRows = sheet2.CreateRow(0);
                    ICell pathCell = startRows.CreateCell(0);
                    pathCell.SetCellValue("PATH");
                    ICell checkCell = startRows.CreateCell(1);
                    checkCell.SetCellValue("Check");

                    // Loop through the data and create cells in the row
                    for (int i = 1; i < rowCount + 1; i++)
                    {
                        // Create a row at the current index
                        IRow row = sheet2.CreateRow(i);

                        // Create cells in the row for each column
                        ICell cell1 = row.CreateCell(0);
                        cell1.SetCellValue(currentPath[i - 1]);

                        ICell cell2 = row.CreateCell(1);
                        cell2.SetCellValue(replaced[i - 1]);

                    }

                    using (FileStream fileStream = new FileStream(newPath, FileMode.Create, FileAccess.Write))
                    {
                        workbook2.Write(fileStream);
                    }
                    Console.WriteLine("Excel file created successfully.");


                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading Excel file: {ex.Message}");
                    
                }
            }
            else
            {
                Console.WriteLine("File not found.");
            }


        }
    }
}