using DAL;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;


namespace ExcelLibrary
{
    public class Class1
    {
        public List<ExcelPosition> HeaderPositions { get; set; }

        public List<Client> Clients { get; set; }
        public void Method()
        {
            HeaderPositions = new List<ExcelPosition> {
                new ExcelPosition{ Text ="ID" },
                new ExcelPosition { Text ="Фамилия имя"},
                new ExcelPosition{ Text ="Дата рождения" },
                new ExcelPosition{ Text ="Номер телефона"},
                new ExcelPosition{Text ="Адрес" },
                new ExcelPosition{ Text = "ИНН"}
            };

            var fileName = "example.xlsx";
            var FilePath = "../../../Template";

            ClientConverter clientConverter = new ClientConverter();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(FilePath + "\\" + fileName, false))
            {   
 
                WorkbookPart workbookPart = doc.WorkbookPart;
                
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                

                ArrayList data = new ArrayList();
                foreach (Row r in sheetData.Elements<Row>())
                {
                    clientConverter.ClientConvert(r, HeaderPositions, workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>());
                }

                foreach (Row r in sheetData.Elements<Row>())
                {
                    int PositionIndex = 0;
                    Client client = new Client();
                    Dictionary<string, string> CustomerDict = new Dictionary<string, string> {
                            { "ID", null },
                            {"Фамилия имя", null},
                            { "Дата рождения" , null},
                            { "Номер телефона", null},
                            {"Адрес", null},
                            { "ИНН", null}
                        };
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        var Coordinate = clientConverter.GetPosition(c.CellReference);
                        
                        if (Coordinate[0] == HeaderPositions[PositionIndex].x && int.Parse(Coordinate[1]) > int.Parse(HeaderPositions[PositionIndex].y))
                        {
                            Worksheet worksheet = new Worksheet();
                            if (c.DataType == null && !string.IsNullOrEmpty(c.InnerText))
                            {
                                CustomerDict[HeaderPositions[PositionIndex].Text] = c.InnerText;
                                PositionIndex++;
                                continue;
                            }    
                            var stringId = Convert.ToInt32(c.InnerText);
                            var Test = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                            CustomerDict[HeaderPositions[PositionIndex].Text] = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                            PositionIndex++;
                        }
                    }
                }
            }
        }
    }

    public class ClientConverter
    {
        public void ClientConvert(Row row, List<ExcelPosition> positions, IEnumerable<SharedStringItem> items )
        {
            int PositionIndex = 0;

            foreach (Cell c in row.Elements<Cell>())
            {
                if (c.DataType != null && (c.DataType == CellValues.String || c.DataType == CellValues.Date || c.DataType == CellValues.Number || c.DataType == CellValues.SharedString))
                {
                    var stringId = Convert.ToInt32(c.InnerText);
                    var Text = items.ElementAt(stringId).InnerText;
                    if (items.ElementAt(stringId).InnerText.Contains(positions[PositionIndex].Text))
                    {
                        var Coordinates = GetPosition(c.CellReference);
                        positions[PositionIndex].x = Coordinates[0];
                        positions[PositionIndex].y = Coordinates[1];
                        PositionIndex++;
                    }
                }

            }
        }

        public string[] GetPosition(string Position)
        {
            var Index = GetAlphabet(Position);
            return new string[] { Position.Substring(0, Index), Position.Substring(Index) };
        }

        private int GetAlphabet(string Text)
        {
            var Sym = Text.First( i => int.TryParse(i.ToString(), out int result));
            var Index = Text.IndexOf(Sym.ToString());
            return Index;
        }

    }


    public class ExcelPosition
    {   
        public string Text { get; set; }

        public string x { get; set; }
        
        public string y { get; set; }
    }


}
