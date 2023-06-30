using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelConverter
{
    internal class Program
    {
        public static string input = "";
        static void Main(string[] args)
        {
            Prompt();
            GetInput();
        }
        public static void Prompt()
        {
            Console.WriteLine("Choose action:");
            Console.WriteLine("1 - Check duplicated keys");
            Console.WriteLine("2 - Copy via key"); //copies lines between sheets basing on "key" column
            Console.WriteLine("3 - Copy via text"); //copies lines between sheets basing on "english" column
            Console.WriteLine("4 - Copy one language"); //copies lines between sheets basing on "english" column
            Console.WriteLine("5 - Check for missing values in collumn"); //copies lines from other excell file, basing on keys
        }
        public static void GetInput()
        {
            Program.input = Console.ReadLine();
            switch(Program.input){
                case "1":
                    CheckKeys();
                    break;
                case "2":
                    CopyViaKey();
                    break;
                case "3":
                    CopyViaText();
                    break;
                case "4":
                    SingleLanguage();
                    break;
                case "5":
                    ColumnCheck();
                    break;
                default:
                    WrongPrompt();
                    break;
            }


        }
        public static void WrongPrompt()
        {
            Console.Clear();
            Prompt();
            Console.WriteLine("Insert proper corresponding number");
            GetInput();
        }
        public static void SingleLanguage()
        {
            string srcPath;
            string destPath;
            int sourceColumn = 8; //column to copy from source doc
            int targetColumn = 8; // collumn to copy to in destination doc

            var keys = new Dictionary<string, string>();
            // only one instance of excel
            Excel.Application excelApplication = new Excel.Application();
            //Opening of first worksheet
            srcPath = "K:\\JkobScrip\\Lang_PTBR.xlsx";
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            destPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);

            int targetRowsAmmount = destworkSheet.UsedRange.Rows.Count;

            int rowsammount = srcworkSheet.UsedRange.Rows.Count;

            //ignore first rows because useless stuff there
            for (int i = 3; i <= rowsammount; i++)
            {
                if (srcworkSheet.Cells[i, 1].Value != null)
                {

                    if (!keys.ContainsKey(srcworkSheet.Cells[i, 1].Value.ToString()))
                    {
                        if (srcworkSheet.Cells[i, sourceColumn].Value != null)
                        {
                            keys.Add(srcworkSheet.Cells[i, 1].Value.ToString(), srcworkSheet.Cells[i, sourceColumn].Value.ToString());
                        }

                    }
                }
            }


            for (int g = 3; g <= targetRowsAmmount; g++)
            {
                if (destworkSheet.Cells[g, 1].Value != null)
                {
                    if (keys.ContainsKey(destworkSheet.Cells[g, 1].Value.ToString()))
                    {
                        if (destworkSheet.Cells[g, targetColumn].Value != keys[destworkSheet.Cells[g, 1].Value.ToString()])
                        {
                            destworkSheet.Cells[g, targetColumn].Value = keys[destworkSheet.Cells[g, 1].Value.ToString()];
                            destworkSheet.Cells[g, targetColumn].Interior.Color = Excel.XlRgbColor.rgbGoldenrod;
                        }
                    }
                }

            }

            Console.WriteLine("Finished");
            //Console.ReadKey();
            excelApplication.Visible = true;

            srcworkBook.Save();
            srcworkBook.Close(true, null, null);
            excelApplication.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                excelApplication = null;
            }
            catch (Exception ex)
            {
                excelApplication = null;
            }
            finally
            {
                GC.Collect();
            }

        }

        public static void ColumnCheck()
        {
            //Color cells without value in corresponding column
            int targetColumn=8;
            string srcPath;
            string destPath;
            var keys = new Dictionary<string, string>();
            // only one instance of excel
            Excel.Application excelApplication = new Excel.Application();
            //Opening of first worksheet
            srcPath = "K:\\JkobScrip\\Lang_PTBR.xlsx";
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            destPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);

            int targetRowsAmmount = destworkSheet.UsedRange.Rows.Count;

            int rowsammount = srcworkSheet.UsedRange.Rows.Count;



            for (int g = 3; g <= targetRowsAmmount; g++)
            {
                if (destworkSheet.Cells[g, 2].Value != null)
                {
                    if (destworkSheet.Cells[g,targetColumn].Value == null)
                    {
                        using (StreamWriter writetext = new StreamWriter("write.txt", true))
                        { 
                            writetext.WriteLine(destworkSheet.Cells[g, 2].Value);
                        }
                        destworkSheet.Cells[g, targetColumn].Interior.Color = Excel.XlRgbColor.rgbRed;
                    }
                }

            }

            Console.WriteLine("Finished");
            Console.ReadKey();
            excelApplication.Visible = true;

            srcworkBook.Save();
            srcworkBook.Close(true, null, null);
            excelApplication.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                excelApplication = null;
            }
            catch (Exception ex)
            {
                excelApplication = null;
            }
            finally
            {
                GC.Collect();
            }

        }
        public static void CopyViaText()
        {
            string srcPath;
            string destPath;
            var keys = new Dictionary<string, string>();
            Excel.Application excelApplication = new Excel.Application();
            srcPath = "K:\\JkobScrip\\LangDesc.xlsx";
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            destPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);
            int _targetRowsAmmount = destworkSheet.UsedRange.Rows.Count;
            int _targetColumnsAmmount = destworkSheet.UsedRange.Columns.Count;

            int rowsammount = srcworkSheet.UsedRange.Rows.Count;
            int columnrange = srcworkSheet.UsedRange.Columns.Count;
            for (int j = 4; j <= columnrange; j++)
            {
                for (int i = 3; i <= rowsammount; i++)
                {
                    if (srcworkSheet.Cells[i, 2].Value != null)
                    {

                        if (!keys.ContainsKey(srcworkSheet.Cells[i, 2].Value.ToString()))
                        {
                            if (srcworkSheet.Cells[i, j].Value == null)
                            {
                                keys.Add(srcworkSheet.Cells[i, 2].Value.ToString(), "");
                            }
                            else
                            {
                                keys.Add(srcworkSheet.Cells[i, 2].Value.ToString(), srcworkSheet.Cells[i, j].Value.ToString());
                            }

                        }
                    }

                }


                for (int g = 3; g <= _targetRowsAmmount; g++)
                {
                    if (destworkSheet.Cells[g, 2].Value != null)
                    {
                        if (keys.ContainsKey(destworkSheet.Cells[g, 2].Value))
                        {
                            if (destworkSheet.Cells[g, j].Value != keys[destworkSheet.Cells[g, 2].Value.ToString()])
                            {
                                destworkSheet.Cells[g, j].Value = keys[destworkSheet.Cells[g, 2].Value.ToString()];
                                destworkSheet.Cells[g, j].Interior.Color = Excel.XlRgbColor.rgbSkyBlue;

                            }
                        }
                    }

                }

                keys.Clear();
            }






            Console.WriteLine("Finished");
            Console.ReadKey();
            excelApplication.Visible = true;

            srcworkBook.Save();
            srcworkBook.Close(true, null, null);
            excelApplication.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                excelApplication = null;
            }
            catch (Exception ex)
            {
                excelApplication = null;
            }
            finally
            {
                GC.Collect();
            }





            //Console.WriteLine("Finished");
            //srcworkBook.Save();
            //destworkBook.Save();
            //srcworkBook.Close(SaveChanges: true);
            //destworkBook.Close(SaveChanges: true);
            //excelApplication.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(srcworkSheet);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(srcworkBook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);

            //OBOSOLITE

        }
        public static void CopyViaKey()
        {
            //When we have a big doc with multiple languages translated
            //NOTE: Używać tylko, jak oba doce mają takie same ułożenie kolumn
            Excel.Application srcxlApp;
            Excel.Range srcrange;
            Excel.Application destxlApp;
            Excel.Range destrange;
            string srcPath;
            string destPath;

            int columnOffset = 4; //number of columns ignored (usually 3: Key, English, Polish)
            var keys = new Dictionary<string, string>();
            // only one instance of excel
            Excel.Application excelApplication = new Excel.Application();
            //Opening of first worksheet
            srcPath = "K:\\JkobScrip\\LangDesc.xlsx";
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            destPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);
            int _targetRowsAmmount = destworkSheet.UsedRange.Rows.Count;
            int _targetColumnsAmmount = destworkSheet.UsedRange.Columns.Count;

            int rowsammount = srcworkSheet.UsedRange.Rows.Count;
            int columnrange = srcworkSheet.UsedRange.Columns.Count;
            for (int j = columnOffset; j <= columnrange; j++)
            {
                //i=3 bo 3 piersze wiersze i tak są useless
                for (int i = 3; i <= rowsammount; i++)
                {
                    if (srcworkSheet.Cells[i, 1].Value != null)
                    {

                        if (!keys.ContainsKey(srcworkSheet.Cells[i, 1].Value.ToString()))
                        {
                            if (srcworkSheet.Cells[i, j].Value == null)
                            {
                                keys.Add(srcworkSheet.Cells[i, 1].Value.ToString(), "");
                            }
                            else
                            {
                                keys.Add(srcworkSheet.Cells[i, 1].Value.ToString(), srcworkSheet.Cells[i, j].Value.ToString());
                            }

                        }
                    }

                }


                for (int g = 3; g <= _targetRowsAmmount; g++)
                {
                    if (destworkSheet.Cells[g, 1].Value != null)
                    {
                        if (keys.ContainsKey(destworkSheet.Cells[g, 1].Value))
                        {
                            if (destworkSheet.Cells[g, j].Value != keys[destworkSheet.Cells[g, 1].Value.ToString()])
                            {
                                destworkSheet.Cells[g, j].Value = keys[destworkSheet.Cells[g, 1].Value.ToString()];
                            }
                        }
                    }

                }

                keys.Clear();
            }



            Console.WriteLine("Finished");
            Console.ReadKey();
            excelApplication.Visible = true;

            srcworkBook.Save();
            srcworkBook.Close(true, null, null);
            excelApplication.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                excelApplication = null;
            }
            catch (Exception ex)
            {
                excelApplication = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void CheckKeys()
        {

            string srcPath;
            string destPath;
            var keys = new Dictionary<string, int>();
            // only one instance of excel
            Excel.Application excelApplication = new Excel.Application();
            //Opening of first worksheet and copying
            srcPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            Excel.Range excelrange = srcworkSheet.UsedRange;
            int rowsammount = srcworkSheet.UsedRange.Rows.Count;
            for (int i = 1; i <= rowsammount; i++)
            {
                if(srcworkSheet.Cells[i, 1].Value != null)
                {

                    if (!keys.ContainsKey(srcworkSheet.Cells[i, 1].Value.ToString()))
                    {
                        keys.Add(srcworkSheet.Cells[i, 1].Value.ToString(), 1);
                    }
                    else 
                    {
                        Console.WriteLine(srcworkSheet.Cells[i, 1].Value.ToString());
                    }

                }
                
            }
            Console.WriteLine(rowsammount - keys.Count);
            Console.ReadLine();


            Console.WriteLine("Finished");
            Console.ReadKey();
            excelApplication.Visible = true;
            srcworkBook.Close(true, null, null);
            excelApplication.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                excelApplication = null;
            }
            catch (Exception ex)
            {
                excelApplication = null;
            }
            finally
            {
                GC.Collect();
            }
        }




    }
}
