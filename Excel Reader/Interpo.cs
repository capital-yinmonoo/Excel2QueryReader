using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab
namespace Excel_Reader
{
    public class Interpo
    {
        public   void getExcelFile(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new  Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("TableName");
            dt.Columns.Add("SheetName");
        
            for  (int j=0; j < xlWorkbook.Sheets.Count; j ++)
            {
              // var f= new Microsoft.CSharp.RuntimeBinder.DynamicMetaObjectProviderDebugView(xlWorkbook.Sheets[1]).Items[32];
                var name = xlWorkbook.Sheets[1];
                dt.Rows.Add(
                    new object[] { 
                     (j+1).ToString(),
                     1,1
                   //  xlWorkbook.Worksheets[j] as 
                    }
                    );
            }

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            //for (int i = 1; i <= rowCount; i++)
            //{
            //    for (int j = 1; j <= colCount; j++)
            //    {
            //        //new line
            //        if (j == 1)
            //            Console.Write("\r\n");

            //        //write the value to the console
            //        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
            //    }
            //} 
            GC.Collect();
            GC.WaitForPendingFinalizers(); 
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet); 
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook); 
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        public DataTable getExcelFile2( string path)
        {

            try
            { 
                Excel.Application oExcel = new Excel.Application(); 
                string filepath =    path; 
                Excel.Workbook WB = oExcel.Workbooks.Open(filepath); 
                string ExcelWorkbookname = WB.Name; 
                int worksheetcount = WB.Worksheets.Count;
                DataTable dt = new DataTable();
                dt.Columns.Add("No");
                dt.Columns.Add("TableName");
                dt.Columns.Add("SheetName");

                for (int j = 0; j < worksheetcount; j++)
                {
                    dt.Rows.Add(
                  new object[] {
                     (j+1).ToString(),
                      "",
                       ((Excel.Worksheet)WB.Worksheets[j+1]).Name
                      //  xlWorkbook.Worksheets[j] as 
                  }
                  );
                }
                return dt;
            }
            catch (Exception ex)
            { 
                string error = ex.Message;
            }
            return null;

        }
    }
}    
