using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXml_Test_2
{
    class ExcelDocument
    {
        public List<Document> Get(string path)
        {
            int[] company_index = new int[7] { 2, 6, 1, 0, 5, 4, 3 };
            List<Document> documents = new List<Document>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;

                Sheets sheets = wbPart.Workbook.Sheets; // It is necessary only for getting names of company.
                List<string> company_names = new List<string>();
                foreach (Sheet sheet in sheets) company_names.Add(sheet.Name);
                

                int comp = 0;
                foreach (WorksheetPart worksheetPart in wbPart.WorksheetParts)
                {
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    List<string> positions = new List<string>();
                    int i = 0;
                    foreach (Row r in sheetData.Elements<Row>())
                    {
                        if (i == 0)
                        {
                            i++;
                            continue;
                        }
                        else if (i == 1)
                        {
                            foreach (Cell cell in r.Elements<Cell>())
                            {
                                positions.Add(CellString(cell, wbPart));
                            }
                            i++;
                            continue;
                        }
                        else if (i > 1)
                        {
                            Document document = new Document();
                            document.Company = company_names[company_index[comp]];
                            int j = 0;
                            foreach (Cell cell in r.Elements<Cell>())
                            {
                                if (j == 0)
                                {
                                    document.Name = CellString(cell, wbPart);
                                    j++;
                                    continue;
                                }
                                else if (CellString(cell, wbPart) != null)
                                {
                                    Position temp = new Position();
                                    temp.Name = positions[FromStringtoInt(cell.CellReference)];
                                    temp.Level = Level(cell);
                                    document.positions.Add(temp);
                                }
                                j++;
                            }
                            documents.Add(document);
                        }
                    }
                    comp++;
                }

            }
            return documents;
        }
        public static string CellString(Cell theCell, WorkbookPart wbPart)
        {
            string value = null;
            if (theCell.InnerText.Length > 0)
            {
                value = theCell.InnerText;

                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }
        public string Level(Cell cell)
        {
            if (cell.StyleIndex.Value <= 30 && cell.StyleIndex.Value >= 25 || cell.StyleIndex.Value == 52 || cell.StyleIndex.Value == 58 ||
                cell.StyleIndex.Value == 59 || cell.StyleIndex.Value == 66 || cell.StyleIndex.Value == 72 || cell.StyleIndex.Value == 73)
                return "Основной";
            return "Общий";
        }
        public int FromStringtoInt(string cellReference)
        {
            string res = cellReference;
            int t = res[0] - 65;
            if (res[1] >= 65 && res[1] <= 90)
            {
                t = (t + 1) * 26 + res[1] - 65;
            }
            return t;
        }
        
    }
}
