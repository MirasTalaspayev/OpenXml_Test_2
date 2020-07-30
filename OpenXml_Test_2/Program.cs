using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml_Test_2
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\Miras Talaspayev\Downloads\Telegram Desktop\BIG_DSM_MTX_004_RU_Матрица_процессов_Холдинга (2).xlsx";
            Test(path);
            
            Console.ReadLine();
        }
        static void Test(string path)
        {
            ExcelDocument one = new ExcelDocument();
            List<Document> res = one.Get(path);
            foreach (Document doc in res)
            {
                Console.WriteLine("Company: {0}.\n\tDocument: {1}", doc.Company, doc.Name);
                foreach (Position pos in doc.positions)
                {
                    Console.WriteLine("\t\tPosition: {0}; Level: {1}", pos.Name, pos.Level);
                }
            }
        }

        static void Test_1()
        {
            string path = @"C:\Users\Miras Talaspayev\Downloads\Telegram Desktop\Test_OpenXml.xlsx"; 
            List<string> company_names = new List<string>();
            int[] company_index = new int[5] { 2, 1, 0, 4, 3 };

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                Sheets sheets = wbPart.Workbook.Sheets; // It is necessary only for getting names of company.
                foreach (Sheet sheet in sheets)
                {
                    company_names.Add(sheet.Name);
                }

                int comp = 0;
                foreach (WorksheetPart worksheetPart in wbPart.WorksheetParts)
                {
                    Console.WriteLine("Sheet: {0}. ", company_names[company_index[comp]]);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    int i = 0;
                    List<string> Positions = new List<string>();

                    foreach (Row rows in sheetData.Elements<Row>())
                    {
                        if (i == 0)
                        {
                            i++;
                            continue;
                        }
                        if (i == 1)
                        {
                            foreach (Cell cell in rows.Elements<Cell>())
                                Positions.Add(ExcelDocument.CellString(cell, wbPart));
                            i++;
                            continue;
                        }
                        int j = 0;
                        string docs = null;
                        
                        foreach (Cell cell in rows.Elements<Cell>())
                        {
                            if (j == 0)
                            {
                                docs = (ExcelDocument.CellString(cell, wbPart));
                                break; 
                            }
                            
                        }
                        Console.WriteLine("\t" + docs);
                        foreach (Cell cell in rows.Elements<Cell>())
                        {

                            if (ExcelDocument.CellString(cell, wbPart) == "+")
                            {

                                Console.WriteLine("\t\t" + Positions[j] + " " + cell.CellReference);

                            }
                            j++;
                        }
                        i++;
                    }
                    Console.WriteLine();
                    comp++;
                }
            }
        }
    }
}

