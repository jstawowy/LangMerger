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
            Console.WriteLine("1 - Check keys without value");
            Console.WriteLine("2 - Copy via key"); //copies lines between sheets basing on "key" column
            Console.WriteLine("3 - Copy via text"); //copies lines between sheets basing on "english" column
            Console.WriteLine("4 - Copy PT language"); //copies lines from other excell file, basing on keys
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
                    Portugal();
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
        public static void Portugal()
        {
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


            for (int i = 3; i <= rowsammount; i++)
            {
                if (srcworkSheet.Cells[i, 1].Value != null)
                {

                    if (!keys.ContainsKey(srcworkSheet.Cells[i, 1].Value.ToString()))
                    {
                        if (srcworkSheet.Cells[i, 8].Value != null)
                        {
                            keys.Add(srcworkSheet.Cells[i, 1].Value.ToString(), srcworkSheet.Cells[i, 8].Value.ToString());
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
                        if (destworkSheet.Cells[g, 8].Value != keys[destworkSheet.Cells[g, 1].Value.ToString()])
                        {
                            destworkSheet.Cells[g, 8].Value = keys[destworkSheet.Cells[g, 1].Value.ToString()];
                            destworkSheet.Cells[g, 8].Interior.Color = Excel.XlRgbColor.rgbGoldenrod;
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

        public static void PortugalCheck()
        {
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
                    if (destworkSheet.Cells[g,8].Value == null)
                    {
                        using (StreamWriter writetext = new StreamWriter("write.txt", true))
                        { 
                            writetext.WriteLine(destworkSheet.Cells[g, 2].Value);
                        }
                        destworkSheet.Cells[g, 8].Interior.Color = Excel.XlRgbColor.rgbRed;
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
        public static void CheckPolish()
        {

            string destPath;
            var keys = new Dictionary<string, string>();
            // only one instance of excel
            Excel.Application excelApplication = new Excel.Application();

            destPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);


            int rowammount = destworkSheet.UsedRange.Rows.Count;

            for (int i = 3; i < rowammount; i++)
            {
                if (destworkSheet.Cells[i, 3].Value == null && destworkSheet.Cells[i, 2].Value != null)
                {
                    destworkSheet.Cells[i, 3].Interior.Color = Excel.XlRgbColor.rgbMediumVioletRed;
                }
            }











            Console.WriteLine("Finished");
            //Console.ReadKey();
            excelApplication.Visible = true;

            destworkBook.Save();
            destworkBook.Close(true, null, null);
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
        public static void CheckString()
        {
            Excel.Application srcxlApp;
            Excel.Range srcrange;
            Excel.Application destxlApp;
            Excel.Range destrange;
            string srcPath;
            string destPath;
            var keys = new Dictionary<string, int>();
            // only one instance of excel
            Excel.Application excelApplication = new Excel.Application();
            //Opening of first worksheet and copying
            srcPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            int rangeRows = srcworkSheet.UsedRange.Rows.Count;
            int rangeColumns = srcworkSheet.UsedRange.Columns.Count;
            string getString;
            char dots = '\u2026';
            string teststring = " ";
            teststring = teststring + dots;


            if (srcworkSheet.Cells[1730, 2].Value != null)
            {
                getString = srcworkSheet.Cells[1730, 2].Value.ToString();
                Console.WriteLine(getString);
                if (getString.Contains(teststring))
                {

                    Console.WriteLine("Old string: " + getString);
                    getString.Replace(teststring, "...");
                    Console.WriteLine("New string: " + getString);




                }
                Console.ReadLine();
            }
            srcworkBook.Save();
            srcworkBook.Close(true);
            excelApplication.Quit();

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
                                //Console.WriteLine(destworkSheet.Cells[g, 2].Value);
                                // Console.WriteLine(destworkSheet.Cells[g, j].Value);

                            }
                        }
                    }

                }

                keys.Clear();
            }







            //destPath = "K:\\JkobScrip\\Lang.xlsx";
            //Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            //Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);
            //int _targetRowsAmmount = destworkSheet.UsedRange.Rows.Count;
            //float _targetRowNumber = 8;
            //for (int i = 3; i <= _targetRowsAmmount; i++)
            //{
            //    if(destworkSheet.Cells[i, 1].Value != null)
            //    {
            //        if (keys.ContainsKey(destworkSheet.Cells[i, 1].Value))
            //        {
            //            if (destworkSheet.Cells[i, 1].Value != keys[destworkSheet.Cells[i, 1].Value.ToString()])
            //            {
            //                destworkSheet.Cells[i, _targetRowNumber].Value = keys[destworkSheet.Cells[i, 1].Value.ToString()];
            //            }
            //        }
            //    }


            //}



            Console.WriteLine("Finished");
            srcworkBook.Save();
            destworkBook.Save();
            srcworkBook.Close(SaveChanges: true);
            destworkBook.Close(SaveChanges: true);
            excelApplication.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(srcworkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(srcworkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);

        }
        public static void CopyViaKey()
        {
            Excel.Application srcxlApp;
            Excel.Range srcrange;
            Excel.Application destxlApp;
            Excel.Range destrange;
            float _rowNumber = 8;
            string srcPath;
            string destPath;
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
            for (int j = 4; j <= columnrange; j++)
            {
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







            //destPath = "K:\\JkobScrip\\Lang.xlsx";
            //Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            //Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);
            //int _targetRowsAmmount = destworkSheet.UsedRange.Rows.Count;
            //float _targetRowNumber = 8;
            //for (int i = 3; i <= _targetRowsAmmount; i++)
            //{
            //    if(destworkSheet.Cells[i, 1].Value != null)
            //    {
            //        if (keys.ContainsKey(destworkSheet.Cells[i, 1].Value))
            //        {
            //            if (destworkSheet.Cells[i, 1].Value != keys[destworkSheet.Cells[i, 1].Value.ToString()])
            //            {
            //                destworkSheet.Cells[i, _targetRowNumber].Value = keys[destworkSheet.Cells[i, 1].Value.ToString()];
            //            }
            //        }
            //    }


            //}



            Console.WriteLine("Finished");
            srcworkBook.Save();
            destworkBook.Save();
            srcworkBook.Close(SaveChanges: true);
            destworkBook.Close(SaveChanges: true);
            excelApplication.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(srcworkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(srcworkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
        }

        public static void PolishLang()
        {
            string srcPath;
            string destPath;
            var keys = new Dictionary<string, string>();
            // only one instance of excel
            Excel.Application excelApplication = new Excel.Application();
            //Opening of first worksheet
            srcPath = "K:\\JkobScrip\\Lang_Dialogues.xlsx";
            Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
            Excel.Worksheet srcworkSheet = srcworkBook.Worksheets.get_Item(1);

            destPath = "K:\\JkobScrip\\Lang.xlsx";
            Excel.Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
            Excel.Worksheet destworkSheet = destworkBook.Worksheets.get_Item(1);

            int targetRowsAmmount = destworkSheet.UsedRange.Rows.Count;

            int rowsammount = srcworkSheet.UsedRange.Rows.Count;


            for (int i = 1; i <= rowsammount; i++)
            {
                if (srcworkSheet.Cells[i, 1].Value != null)
                {

                    if (!keys.ContainsKey(srcworkSheet.Cells[i, 1].Value.ToString()))
                    {
                        if (srcworkSheet.Cells[i, 3].Value != null)
                        {
                            keys.Add(srcworkSheet.Cells[i, 1].Value.ToString(), srcworkSheet.Cells[i, 3].Value.ToString());
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
                        if (destworkSheet.Cells[g, 3].Value != keys[destworkSheet.Cells[g, 1].Value.ToString()])
                        {
                            destworkSheet.Cells[g, 3].Value = keys[destworkSheet.Cells[g, 1].Value.ToString()];
                            destworkSheet.Cells[g, 3].Interior.Color = Excel.XlRgbColor.rgbGoldenrod;
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

            string test = srcworkSheet.Cells[1, 1].Value.ToString();
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
            srcworkBook.Close(false, null, null);
            excelApplication.Quit();
        }


    }
}
