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
using System.Diagnostics;

using System.Reflection;
using System.Data.SqlClient;
namespace RemoteADM
{
    class ExcelFile
    {
        private string excelFilePath = string.Empty;
        public  int rowNumber = 4; // define first row number to enter data in excel

        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet myExcelWorkSheet;

        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }
        void setBorder()
        {

        }
        void setonCreate()
        {

        }

        public int Rownumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }

        public async void openExcel()
        {



        try
            {
                Process[] pro = Process.GetProcessesByName("excel");

                pro[0].Kill();
                pro[0].WaitForExit();



            }
            catch(Exception E)
            {

            }

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
        private void AllBorders(Excel.Range oRange)
        {
            oRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            oRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            oRange.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            oRange.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
            oRange.Borders.Color = Color.Black;
        }

        public async void addDataToExcel(string Name, String Zone, string Entry_Time, string Entry_Date, string Exit_Time, string Exit_Date,int x)
        {

            Excel.Range oRange= myExcelWorkSheet.get_Range("A5", "H"+rowNumber);
            AllBorders(oRange);

            myExcelWorkSheet.Cells[rowNumber, 1] = rowNumber-5;
            myExcelWorkSheet.Cells[rowNumber, 2] = Name;
            myExcelWorkSheet.Cells[rowNumber, 3] = Zone;
            myExcelWorkSheet.Cells[rowNumber, 4] = Entry_Time;
            myExcelWorkSheet.Cells[rowNumber, 5] = Entry_Date;
            myExcelWorkSheet.Cells[rowNumber, 6] = Exit_Time;
            myExcelWorkSheet.Cells[rowNumber, 7] = Exit_Date;

            // setting borders

            if (x == 1)
            {
                Find_TimeDiff(myExcelWorkSheet);
            }
            else
            {
                myExcelWorkSheet.Cells[rowNumber, 8] = "00:00";
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

            xlWorkSheet.Cells[2, 1] = "RUNNING ROOM OCCUPANCY STATUS";
            xlWorkSheet.get_Range("A2", "A2").Cells.Font.Size =15;
            xlWorkSheet.get_Range("A2", "A2").Cells.Font.Bold = true;
            Excel.Range oRange = xlWorkSheet.get_Range("A2","F2");
            AllBorders(oRange);
            xlWorkSheet.get_Range("A2","A2").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;




            xlWorkSheet.get_Range("A2:F2", Type.Missing).Merge(Type.Missing);
          
            xlWorkSheet.get_Range("O1:P1", Type.Missing).Merge(Type.Missing);
            xlWorkSheet.get_Range("O2:S2", Type.Missing).Merge(Type.Missing);


            setBorder();
           

            xlWorkSheet.Cells[5, 1] = "Sr No.";
            xlWorkSheet.get_Range("A5", "A5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[1].ColumnWidth = 6;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; 


            xlWorkSheet.Cells[5, 2] = "Name";
            xlWorkSheet.get_Range("B5", "B5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[2].ColumnWidth = 25;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[5, 3] = "Zone";
            xlWorkSheet.get_Range("C5", "C5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[3].ColumnWidth = 5;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[5, 4] = "Entry Time";
            xlWorkSheet.get_Range("D5", "D5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[4].ColumnWidth = 9;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[5, 5] = "Entry Date";
            xlWorkSheet.get_Range("E5", "E5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[5].ColumnWidth = 9;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[5, 6] = "Exit Time";
            xlWorkSheet.get_Range("F5", "F5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[6].ColumnWidth = 9;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[5, 7] = "Exit Date";
            xlWorkSheet.get_Range("G5", "G5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[7].ColumnWidth = 9;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[5, 8] = "Duration";
            xlWorkSheet.get_Range("H5", "H5").Cells.Font.Bold = true;
            xlWorkSheet.Columns[8].ColumnWidth = 11;
            xlWorkSheet.get_Range("A5", "A5").Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;


            xlWorkSheet.Cells[1, 7] = "Date:";
            //xlWorkSheet.Cells[2, 7] = "Place:";
            xlWorkSheet.Cells[1, 8] = dateTime.ToString("dd/mm/yyyy");
            //////////////////////Change place value here
           // xlWorkSheet.Cells[2, 8] = "Place:";



            Excel.Range BorderRAnge;
            BorderRAnge = xlWorkSheet.get_Range("A" + rowNumber, "H" + rowNumber);



            // xlWorkSheet.get_Range("C2", "C2").Cells.Font.Size = 15;
            // xlWorkSheet.get_Range("C2", "C2").Cells.Font.Bold = true;
            // xlWorkSheet.Cells[4, 2] = "Total No. of Hours rooms,oblique bed occupied:";

            // xlWorkSheet.get_Range("B5:F5", Type.Missing).Merge(Type.Missing);
            // xlWorkSheet.get_Range("C3", "C3").Cells.Font.Size = 15;
            // xlWorkSheet.get_Range("C3", "C3").Cells.Font.Bold = true;
            // xlWorkSheet.Cells[5, 2] = "Effective Utilization of beds Overall:";


            // try
            // {
            //     xlWorkSheet.Cells[4, 7].NumberFormat = "hh:mm";

            //     xlWorkSheet.Cells[5, 7].NumberFormat = "0.00%";
            //     xlWorkSheet.Cells[5, 7].Formula = "=(((G4*1440)/(108*60*24))*100)";

            // }
            //catch(Exception e)
            // {
            //     MessageBox.Show(e.ToString());
            // }


            //Here saving the file in xlsx
            xlWorkBook.SaveAs(excelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();





             rowNumber = 4;

        Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }
        void Find_TimeDiff(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet)
        {
            try
            {
              
                myExcelWorkSheet.Cells[ rowNumber,8].NumberFormat = "hh:mm";
                myExcelWorkSheet.Cells[rowNumber, 8].Formula = "=(G"+Rownumber+"+F"+Rownumber+")-(E"+Rownumber+"+D"+Rownumber+")";
                //xlWorkSheet.Cells[4, 7].Formula = "=sum(G" + 8 + ":G" + (Rownumber) + ")";
             
            }
            catch(Exception E)
            {
                MessageBox.Show(E.ToString());
            }

        }
    }
}
