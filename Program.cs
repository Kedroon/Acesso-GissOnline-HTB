using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsolidarPlanilhaShippingPlan
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo excelDirectory = new DirectoryInfo(@"C:\Users\migue\OneDrive\Documentos\Visual Studio 2015\Projects\ConsolidarPlanilhaShippingPlan\ConsolidarPlanilhaShippingPlan\Planilhas");
            FileInfo[] excelFiles = excelDirectory.GetFiles("0*");
            foreach (var item in excelFiles)
            {
                Console.WriteLine(item.ToString());
            }
            string workbookPathRede = excelDirectory + @"\" + excelFiles[4].ToString();
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            string workbookPath = workbookPathRede;
            var workbooks = excelApp.Workbooks;
            Excel.Workbook excelWorkbook = workbooks.Open(workbookPath,
                    false, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            excelWorkbook.Application.DisplayAlerts = false;
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            Excel.Worksheet Worksheet = (Excel.Worksheet)excelSheets.get_Item(1);

   
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = excelApp.Workbooks.Add(misValue);
            Excel.Sheets excelSheetsDestino = xlWorkBook.Worksheets;
            Excel.Worksheet WorksheetConsolidado = (Excel.Worksheet)excelSheetsDestino.get_Item(1);
            
            int linhaPO = 1;
            int linhacabecalho;
            while (Worksheet.Cells[linhaPO,1].Value != "PO")
            {
                linhaPO++;
            }
            linhacabecalho = linhaPO;
            linhaPO++;
            int countnull = linhaPO;
            while (Worksheet.Cells[countnull, 6].Value != null)
            {
                countnull++;
            }
            countnull--;
            DateTime data1 = new DateTime(2016, 08, 1);
            DateTime data2 = new DateTime(2016, 08, 19);
            int countETD = 0;
            if (Worksheet.Range["AH" + linhacabecalho].Value.ToString().IndexOf("ETD") != -1)
            {
                countETD++;
            }
            if (Worksheet.Range["AI" + linhacabecalho].Value.ToString().IndexOf("ETD") != -1)
            {
                countETD++;
            }
            if (Worksheet.Range["AJ" + linhacabecalho].Value.ToString().IndexOf("ETD") != -1)
            {
                countETD++;
            }

            Console.WriteLine(countETD);

            int countTerminoETD = -1;

            int linhaDestiny = 1;
            do
            {
                countTerminoETD++;
                int countTermino = linhaPO;

                while (countnull != countTermino)
                {
                    string colunaETD="";
                    try
                    {
                        if (countTerminoETD==0)
                        {
                            colunaETD = "AG";
                        }
                        else if (countTerminoETD == 1)
                        {
                            colunaETD = "AH";
                        }
                        else if (countTerminoETD == 2)
                        {
                            colunaETD = "AI";
                        }
                        else if (countTerminoETD == 3)
                        {
                            colunaETD = "AJ";
                        }

                        string datas = Worksheet.Range[colunaETD + countTermino].Value.ToString();

                        datas = datas.Substring(0, 9);
                        DateTime data = DateTime.Parse(datas);
                        if (data >= data1 && data <= data2)
                        {
                            Console.WriteLine(data);3
                            Excel.Range from = Worksheet.Range["A" + countTermino + ":AG" + countTermino];
                            Console.WriteLine(countTermino);
                            Excel.Range to = WorksheetConsolidado.Range["A" + linhaDestiny + ":AG" + linhaDestiny];

                            from.Copy(Type.Missing);
                            to.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                            if (Worksheet.Range["AH" + linhacabecalho].Value.ToString().IndexOf("ETD") != -1)
                            {
                                Excel.Range from1 = Worksheet.Range["AH" + countTermino + ":AH" + countTermino];
                                Excel.Range to1 = WorksheetConsolidado.Range["AH" + linhaDestiny + ":AH" + linhaDestiny];
                                from1.Copy(Type.Missing);
                                to1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                            if (Worksheet.Range["AI" + linhacabecalho].Value.ToString().IndexOf("ETD") != -1)
                            {
                                Excel.Range from1 = Worksheet.Range["AI" + countTermino + ":AI" + countTermino];
                                Excel.Range to1 = WorksheetConsolidado.Range["AI" + linhaDestiny + ":AI" + linhaDestiny];
                                from1.Copy(Type.Missing);
                                to1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                            if (Worksheet.Range["AJ" + linhacabecalho].Value.ToString().IndexOf("ETD") != -1)
                            {
                                Excel.Range from1 = Worksheet.Range["AJ" + countTermino + ":AJ" + countTermino];
                                Excel.Range to1 = WorksheetConsolidado.Range["AJ" + linhaDestiny + ":AJ" + linhaDestiny];
                                from1.Copy(Type.Missing);
                                to1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                            if (countETD == 0)
                            {
                                Excel.Range from1 = Worksheet.Range["AH" + countTermino + ":AR" + countTermino];
                                Excel.Range to1 = WorksheetConsolidado.Range["AK" + linhaDestiny + ":AX" + linhaDestiny];
                                from1.Copy(Type.Missing);
                                to1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                            if (countETD == 1)
                            {
                                Excel.Range from1 = Worksheet.Range["AI" + countTermino + ":AT" + countTermino];
                                Excel.Range to1 = WorksheetConsolidado.Range["AK" + linhaDestiny + ":AX" + linhaDestiny];
                                from1.Copy(Type.Missing);
                                to1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                            if (countETD == 2)
                            {
                                Excel.Range from1 = Worksheet.Range["AJ" + countTermino + ":AV" + countTermino];
                                Excel.Range to1 = WorksheetConsolidado.Range["AK" + linhaDestiny + ":AX" + linhaDestiny];
                                from1.Copy(Type.Missing);
                                to1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                            if (countETD == 3)
                            {
                                Excel.Range from1 = Worksheet.Range["AK" + countTermino + ":AX" + countTermino];
                                Excel.Range to1 = WorksheetConsolidado.Range["AK" + linhaDestiny + ":AX" + linhaDestiny];
                                from1.Copy(Type.Missing);
                                to1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                            
                            linhaDestiny++;
                        }
                    }
                    catch (Exception err)
                    {
                        Console.WriteLine(err.Message);
                    }
                    countTermino++;
                }
                
            } while (countETD != countTerminoETD);

           

            /*Console.Read();
            Excel.Range from = Worksheet.Range["A"+linhaPO+ ":AG" + countnull];
            Excel.Range to = WorksheetConsolidado.Range["A1:AG" +countnull];
            from.Copy(Type.Missing);
            to.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            */
            Console.Read();
            if (Worksheet.Cells[linhaPO,33].Value.IndexOf("ETD")== -1)
            {
                Console.WriteLine("oi");

            }
            else
            {
                Console.WriteLine("tchau");
            }
            
            try
            {
                Worksheet.Range["AK:AK"].Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                
            }
            
            
        }
    }
}
