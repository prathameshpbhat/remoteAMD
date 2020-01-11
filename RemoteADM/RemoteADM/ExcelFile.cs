using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Reflection;
using System.Data.SqlClient;
namespace RemoteADM
{
    class ExcelFile
    {
        private string excelFilePath = string.Empty;
        public  int rowNumber = 8; // define first row number to enter data in excel

        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet myExcelWorkSheet;

        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }

        public int Rownumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }

        public async void openExcel()
        {
            myExcelApplication = null;

            myExcelApplication = new Excel.Application(); // create Excell App
            myExcelApplication.DisplayAlerts = false; // turn off alerts


            myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks._Open(excelFilePath, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

            int numberOfWorkbooks = myExcelApplication.Workbooks.Count; // get number of workbooks (optional)

            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1]; // define in which worksheet, do you want to add data
            myExcelWorkSheet.Name = "WorkSheet 1"; // define a name for the worksheet (optinal)

            int numberOfSheets = myExcelWorkbook.Worksheets.Count; // get number of worksheets (optional)
        }

        public async void addDataToExcel(string Name, String Zone, string Entry_Time, string Entry_Date, string Exit_Time, string Exit_Date,int x)
        {
            
            myExcelWorkSheet.Cells[rowNumber, 1] = Name;
            myExcelWorkSheet.Cells[rowNumber, 2] = Zone;
            myExcelWorkSheet.Cells[rowNumber, 3] = Entry_Time;
            myExcelWorkSheet.Cells[rowNumber, 4] = Entry_Date;
            myExcelWorkSheet.Cells[rowNumber, 5] = Exit_Time;
            myExcelWorkSheet.Cells[rowNumber, 6] = Exit_Date;
            if(x==1)
            {
                Find_TimeDiff(myExcelWorkSheet);
            }
            else
            {
                myExcelWorkSheet.Cells[rowNumber, 7] = "00:00";
            }
            rowNumber++;  // if you put this method inside a loop, you should increase rownumber by one or wat ever is your logic

        }

        public async void closeExcel()
        {
            try
            {
                myExcelWorkbook.SaveAs(excelFilePath, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value); // Save data in excel


                myExcelWorkbook.Close(true, excelFilePath, System.Reflection.Missing.Value); // close the worksheet


            }
            catch
            {
              
            }
            finally
            {
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                }
            }
            Marshal.ReleaseComObject(myExcelWorkSheet);
            Marshal.ReleaseComObject(myExcelWorkbook);
            Marshal.ReleaseComObject(myExcelApplication);

        }
       public async void create_newExcel()
        {
            DateTime dateTime = DateTime.UtcNow.Date;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
              
                return;
            }


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[2, 3] = "RUNNING ROOM OCCUPANCY STATUS";
            xlWorkSheet.get_Range("C2", "C2").Cells.Font.Size = 20;
            xlWorkSheet.get_Range("C2", "C2").Cells.Font.Bold = true;



            xlWorkSheet.get_Range("C1:M1", Type.Missing).Merge(Type.Missing);
            xlWorkSheet.get_Range("C2:M2", Type.Missing).Merge(Type.Missing);
            xlWorkSheet.get_Range("C3:M3", Type.Missing).Merge(Type.Missing);
            xlWorkSheet.get_Range("O1:P1", Type.Missing).Merge(Type.Missing);
            xlWorkSheet.get_Range("O2:S2", Type.Missing).Merge(Type.Missing);

            xlWorkSheet.get_Range("B4:F4", Type.Missing).Merge(Type.Missing);
           
            

            xlWorkSheet.Cells[7, 1] = "Name";
            xlWorkSheet.Cells[7, 2] = "Zone";
            xlWorkSheet.Cells[7, 3] = "Entry Time";
            xlWorkSheet.Cells[7, 4] = "Entry Date";
            xlWorkSheet.Cells[7, 5] = "Exit Time";
            xlWorkSheet.Cells[7, 6] = "Exit Date";
            xlWorkSheet.Cells[7, 7] = "Duration";


            xlWorkSheet.Cells[1, 14] = "Date:";
            xlWorkSheet.Cells[2, 14] = "Place:";
            xlWorkSheet.Cells[1, 15] = dateTime.ToString("dd/mm/yyyy");
            //////////////////////Change place value here
            xlWorkSheet.Cells[2, 15] = "Place:";


            xlWorkSheet.get_Range("C2", "C2").Cells.Font.Size = 15;
            xlWorkSheet.get_Range("C2", "C2").Cells.Font.Bold = true;
            xlWorkSheet.Cells[4, 2] = "Total No. of Hours rooms,oblique bed occupied:";

            xlWorkSheet.get_Range("B5:F5", Type.Missing).Merge(Type.Missing);
            xlWorkSheet.get_Range("C3", "C3").Cells.Font.Size = 15;
            xlWorkSheet.get_Range("C3", "C3").Cells.Font.Bold = true;
            xlWorkSheet.Cells[5, 2] = "Effective Utilization of beds Overall:";


            try
            {
                xlWorkSheet.Cells[4, 7].NumberFormat = "hh:mm";

                xlWorkSheet.Cells[5, 7].NumberFormat = "0.00%";
                xlWorkSheet.Cells[5, 7].Formula = "=(((G4*1440)/(108*60*24))*100)";

            }
           catch(Exception e)
            {
                MessageBox.Show(e.ToString());
            }


            //Here saving the file in xlsx
            xlWorkBook.SaveAs(excelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();





             rowNumber = 8;

        Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }
        void Find_TimeDiff(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet)
        {
            try
            {
              
                myExcelWorkSheet.Cells[ rowNumber,7].NumberFormat = "hh:mm";
                myExcelWorkSheet.Cells[rowNumber, 7].Formula = "=(F"+Rownumber+"+E"+Rownumber+")-(D"+Rownumber+"+C"+Rownumber+")";
                xlWorkSheet.Cells[4, 7].Formula = "=sum(G" + 8 + ":G" + (Rownumber) + ")";
             
            }
            catch(Exception E)
            {
                MessageBox.Show(E.ToString());
            }

        }
    }
}
