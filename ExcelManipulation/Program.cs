// Created By Said Ridaoui

namespace ExcelManipulation
{
    using ClosedXML.Excel;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.IO.Packaging;
    using System.Linq;
    using System.Net.WebSockets;
    using System.Reflection;
    using System.Text;  
    using System.Threading.Tasks;
    using System.Windows.Forms;

    /// <summary>
    /// Defines the <see cref="Program" />.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// The Main.
        /// </summary>
       
        [STAThread]
        internal static void Main(string[] args)
        {
            MessageBox.Show("Bitte laden Sie die Excel-Datei hoch, um sie zu transponieren");
            
            uploadFile();
        }

        /// <summary>
        /// The getExtention Function to get the extention of the File based on filePath.
        /// </summary>
        /// <param name="filePath">The filePath<see cref="string"/>.</param>
        /// <returns>Extention <see cref="string"/></returns>
        public static string getExtention(string filePath)        {
            return Path.GetExtension(filePath);
        }

        /// <summary>
        /// The displayTranspose Function to display the transposed Excel.
        /// </summary>
        /// <param name="transposedPath">the Transposed Path<see cref="string"/>.</param>
        public static void displayTranspose(string transposedPath)        {
            
                
            //Started reading the Excel file.
            using (var workBook = new XLWorkbook(transposedPath))
             {
                 var workSheet = workBook.Worksheet(1);
                 var firstRowUsed = workSheet.FirstRowUsed();
                 var firstPossibleAddress = workSheet.Row(firstRowUsed.RowNumber()).FirstCell().Address;
                 var lastPossibleAddress = workSheet.LastCellUsed().Address;
  
                 // Getting a range with the remainder of the worksheet data / the range used)
                 var range = workSheet.Range(firstPossibleAddress, lastPossibleAddress).AsRange(); //.RangeUsed();

                 // Treat the range as a table to be able to use the column names
                 var table = range.AsTable();
  
                 //putting all the Columns from the Tranposed Excel into A List called dataList
                 var dataList = new List<string[]>
                 {
                     table.DataRange.Rows()
                         .Select(tableRow =>
                             tableRow.Field("Kundenname")
                                 .GetString())
                         .ToArray(),

                     table.DataRange.Rows()
                         .Select(tableRow => tableRow.Field("Kundennummer").GetString())
                         .ToArray(),

                     table.DataRange.Rows()
                     .Select(tableRow => tableRow.Field("Kontakadresse").GetString())
                     .ToArray(),

                     table.DataRange.Rows()
                     .Select(tableRow => tableRow.Field("Telefonnummer").GetString())
                     .ToArray(),

                     table.DataRange.Rows()
                     .Select(tableRow => tableRow.Field("Lieferanschrift").GetString())
                     .ToArray(),

                     table.DataRange.Rows()
                     .Select(tableRow => tableRow.Field("Kommentar 1").GetString())
                     .ToArray(),

                     table.DataRange.Rows()
                     .Select(tableRow => tableRow.Field("Kommentar 2").GetString())
                     .ToArray(),

                 };

                 //Convert the above List to DataTable via the Function the Explicit function ConvertListToDataTable
                 var dataTable = ConvertListToDataTable(dataList);

                 //To avoid removing duplication colum To get unique column values KundenName as Exemple
                 var uniqueCols = dataTable.DefaultView.ToTable(true, "Kundenname");
  
                 //Removing any Empty Rows 
                 for (var i = uniqueCols.Rows.Count - 1; i >= 0; i--)
                 {
                     var dr = uniqueCols.Rows[i];
                     if (dr != null && ((string)dr["Kundenname"] == "None" )){dr.Delete();}
                         
                 }
                 
               
               
             }
        }

        /// <summary>
        /// The ConvertListToDataTable function which take a List as parametre and and return it as Table.
        /// </summary>
        /// <param name="list">The list<see cref="IReadOnlyList{string[]}"/>.</param>
        /// <returns>DataTable<see cref="DataTable"/>.</returns>
        private static DataTable ConvertListToDataTable(IReadOnlyList<string[]> list)
        { 
            // creating the object table where to stock the List as table called filename_modified
             var table = new DataTable("Testdatei_modified");
             var rows = list.Select(array => array.Length).Concat(new[] { 0 }).Max();

             // adding the collumns into the table filename_modified
             table.Columns.Add("Kundenname");
             table.Columns.Add("Kundennummer");
             table.Columns.Add("Kontakadresse");
             table.Columns.Add("Telefonnummer");
             table.Columns.Add("Lieferanschrift");
             table.Columns.Add("Kommentar 1");
             table.Columns.Add("Kommentar 2");

             // inserting the Data of list into the table
             for (var j = 0; j < rows; j++)
             {
                 var row = table.NewRow();
                 row["Kundenname"] = list[0][j];
                 row["Kundennummer"] = list[1][j];
                 row["Kontakadresse"] = list[2][j];
                 row["Telefonnummer"] = list[3][j];
                 row["Lieferanschrift"] = list[4][j];
                 row["Kommentar 1"] = list[5][j];
                 row["Kommentar 2"] = list[6][j];
                 
                 table.Rows.Add(row);
             }
             return table;
        }

        /// <summary>
        /// The excelTranspose function to transpose the excel file Via ClosedXML.
        /// </summary>
        /// <param name="filePath">The filePath parametre<see cref="string"/>.</param>
        public static void excelTranspose(string filePath)        {

            // opening the file via workbook which inherit from the class XLWorkbook
            using (XLWorkbook workbook = new XLWorkbook(filePath)) 
        {   
            // getting the Directory name based on the file path 
            // so that we can store the file in the current directory of the file uploaded
            string folder = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileName(filePath);           
            string pathToCheck =  folder + @"\" + fileName ;
                
            // creating the worksheet to treat the file.
            var ws = workbook.Worksheet(1);
            // styling the transposed file by adding color and changing the font size...
            ws.Range("A1:A7").Rows().Style.Fill.BackgroundColor = XLColor.Orange;
            ws.Range("A1:A7").Style.Font.Bold = true;
            ws.Range("A1:A7").Style.Font.FontSize = 12; 
            ws.Range("A1:A7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;    

            // A1: First used row
            // D7: last used row
            var tableRng = ws.Range("A1:D7");
            tableRng.Transpose(XLTransposeOptions.MoveCells);
            ws.Columns().AdjustToContents();

            string tempfileName = "";
            
        // a Check to see if a file already exists with the same name as the file to upload.
        
        if (System.IO.File.Exists(pathToCheck)) 
        {
          
          while (System.IO.File.Exists(pathToCheck))
          {
            // if a file with this Name already exists,
            // add "_modified" at end of the file Name .

            tempfileName =  Path.GetFileNameWithoutExtension(filePath) + "_modified";
            pathToCheck = folder + @"\" + tempfileName;
 
            
          }

          fileName = tempfileName;
          
          // Notify the user that the file name was changed.
          MessageBox.Show(" A file with the same name already exists " + Environment.NewLine + " Your file was saved as: " + fileName);
        }

        
        

        // Append the name of the file to upload to the path.
        folder =  folder + @"\" + fileName + ".xlsx" ;
            
        // Call the SaveAs method to save the uploaded file to the specified directory.
        workbook.SaveAs(folder);

        // Notify the user that the file was saved successfully.

            string title = "Your Excel file was uploaded successfully.!!!"; 
            string message = "Do you want to Display the Transpose Excel";  
            
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;  
            DialogResult result = MessageBox.Show(message, title, buttons);  
            if (result == DialogResult.Yes) {  

                
                   // calling the fucntion when the user want to display the transposed file
                    displayTranspose(folder);

            } else {  

                    Application.Exit();
                  
            } 

         // Dispose or stop the excel process running after saving the transposed file
        workbook.Dispose();




            

            }
        }

        /// <summary>
        /// The uploadFile function used for uploading the file and to validate the extention.
        /// </summary>
       
        public static void uploadFile(){
            
            var filePath = string.Empty;
            var result = DialogResult.Retry;
            var repeat = true;
            

    using (OpenFileDialog openFileDialog = new OpenFileDialog())
    {       
            // used to take the user directly to file which we want to transpose
            openFileDialog.InitialDirectory = @"ExcelManipulation\file_path";
            
            openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

           // An unfinity Loop until the user choose a valid excel
           while(repeat){

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {  
                        filePath = openFileDialog.FileName;
                        // used to avoid empty file with 0 ko
                        long filelength = new System.IO.FileInfo(filePath).Length;
                        
                        // when the extention is valid and file not empty exit the loop and transpose the file
                        if (getExtention(filePath) == ".xlsx" && filelength >0 ) { excelTranspose(filePath); repeat= false;}


                        // keep showing dialog Message until user upload valid excel
                        else
                        {
                             result = MessageBox.Show(@"Please select a Non Empty Excel Format instead!", @"Please select a valid file..",
                                 MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);


                            if (result == DialogResult.Retry && getExtention(filePath) == ".xlsx"){ excelTranspose(filePath); repeat=false;}
                            if (result == DialogResult.Abort || result == DialogResult.Cancel) { break; }

                            
                        }

   
                }
                // when user tap cancel from the beginning and want to exit
                else { break;}
                        


           }

    }


            
        }
    }
}
