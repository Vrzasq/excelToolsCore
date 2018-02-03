using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace excelToolsCore.utilities
{
    public class Helpers
    {

        /// <summary>
        /// Return cleaned and trimed cell value, if value is null or empty returns string.empty
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="header">column</param>
        /// <param name="data">2D data array</param>
        /// <returns></returns>
        public static string GetValueFromCellCleaned(object[,] data, int row, int col)
        {
            if (data[row, col] == null)
                return string.Empty;
            return CleanReferance(data[row, col].ToString());
        }

        public static string GetValueFromCellAsciiPrintable(object[,] data, int row, int col)
        {
            if (data[row, col] == null)
                return string.Empty;
            return AsciiPrintable(data[row, col].ToString()).Trim();
        }

        public static string GetValueFromCellTrimed(object[,] data, int row, int col)
        {
            if (data[row, col] == null)
                return string.Empty;
            return data[row, col].ToString().Trim();
        }

        public static string CleanFirstAndLastChar(object[,] data, int row, int col)
        {
            string cellValue = GetValueFromCellTrimed(data, row, col);
            if (string.IsNullOrEmpty(cellValue))
                return cellValue;

            StringBuilder sb = new StringBuilder();
            sb.Append(CleanReferance(cellValue[0].ToString()));

            for (int i = 1; i < cellValue.Length - 1; i++)
                sb.Append(cellValue[i]);

            sb.Append(CleanReferance(cellValue[cellValue.Length - 1].ToString()));

            return sb.ToString();
        }


        public static List<string> GetCellWordsCleaned(object[,] data, int row, int col)
        {
            List<string> l = new List<string>();
            if (data[row, col] == null)
                return l;
            string[] ar = data[row, row].ToString().Split(' ');
            for (int i = 0; i < ar.Length; i++)
            {
                string temp = CleanReferance(ar[i]);
                if (temp != string.Empty)
                    l.Add(temp);
            }
            return l;
        }

        public static List<string> GetCellWordsCleanedLowerCase(object[,] data, int row, int col)
        {
            List<string> l = new List<string>();
            if (data[row, col] == null)
                return l;
            string[] ar = data[row, row].ToString().ToLower().Split(' ');
            for (int i = 0; i < ar.Length; i++)
            {
                string temp = CleanReferance(ar[i]);
                if (temp != string.Empty)
                    l.Add(temp);
            }
            return l;
        }

        public static List<string> GetCellWordsTrimedLowerCase(object[,] data, int row, int col)
        {
            List<string> l = new List<string>();
            if (data[row, col] == null)
                return l;
            string[] ar = data[row, row].ToString().ToLower().Split(' ');
            for (int i = 0; i < ar.Length; i++)
            {
                string temp = ar[i];
                if (temp != string.Empty)
                    l.Add(temp.Trim());
            }
            return l;
        }

        public static List<string> GetCellWordsLowerCase(object[,] data, int row, int col)
        {
            List<string> l = new List<string>();
            if (data[row, col] == null)
                return l;
            string[] ar = data[row, row].ToString().ToLower().Split(' ');
            for (int i = 0; i < ar.Length; i++)
            {
                if (ar[i] != string.Empty)
                    l.Add(ar[i]);
            }
            return l;
        }

        /// <summary>
        /// Remove all non alpha numeric characters from string
        /// </summary>
        /// <param name="referance"></param>
        /// <returns></returns>
        public static string CleanReferance(string referance)
        {
            Regex regex = new Regex("[^a-zA-Z0-9 -]");
            return regex.Replace(referance, string.Empty).Trim();
        }


        // all ascii range 0000-007F
        public static string AsciiPrintable(string referance)
        {
            Regex regex = new Regex(@"[^\u0020-\u007E]+");
            return regex.Replace(referance, string.Empty);
        }

        public static string[][] ReadeCsv(string filePath)
        {
            try
            {
                return File.ReadAllLines(filePath).Select(s => s.Split(';')).ToArray();
            }
            catch (IOException e)
            {
                throw e;
            }
        }

        public static FileInfo ConvertToXLSX(FileInfo file)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            Excel.Workbooks temp = excelApp.Workbooks;
            Excel.Workbook wrk;
            try
            {
                wrk = temp.Open(file.FullName);
            }
            catch (IOException e)
            {
                throw e;
            }

            wrk.SaveAs(Filename: file.Name + "x", FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);
            FileInfo newfile = new FileInfo(wrk.FullName);
            excelApp.DisplayAlerts = true;

            wrk.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(wrk);
            Marshal.ReleaseComObject(temp);
            Marshal.ReleaseComObject(excelApp);

            return newfile;
        }

        public static void SaveWithExcel(string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            Excel.Workbooks temp = excelApp.Workbooks;
            Excel.Workbook wrk;

            try
            {
                wrk = temp.Open(filePath);
            }
            catch (IOException e)
            {
                throw e;
            }

            wrk.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);

            excelApp.DisplayAlerts = true;
            wrk.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(wrk);
            Marshal.ReleaseComObject(temp);
            Marshal.ReleaseComObject(excelApp);
        }
    }
}