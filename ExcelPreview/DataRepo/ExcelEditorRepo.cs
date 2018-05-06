using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;

namespace ExcelPreview.DataRepo
{
    public class ExcelEditorRepo
    {
        public void CreateExcelDoc(string fileName, List<Models.CandidateInfo> candidateList)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();
                
                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                // Constructing header
                Row row = new Row();
                foreach (PropertyInfo pi in candidateList[0].GetType().GetProperties())
                {
                    row.Append(ConstructCell(pi.Name, CellValues.String));
                }

                //row.Append(
                //    ConstructCell("CandidateId", CellValues.String),
                //    ConstructCell("CandidateName", CellValues.String),
                //    ConstructCell("Exp", CellValues.String),
                //    ConstructCell("Salary", CellValues.String),
                //    ConstructCell("DOB", CellValues.String));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                foreach (var item in candidateList)
                {
                    row = new Row();

                    row.Append(
                        ConstructCell(item.CandidateId.ToString(), CellValues.Number),
                        ConstructCell(item.CandidateName.ToString(), CellValues.String),
                        ConstructCell(item.Exp.ToString(), CellValues.String),
                        ConstructCell(item.Salary.ToString(), CellValues.String),
                        ConstructCell(item.DOB.ToString("dd/MM/yyyy"), CellValues.String));                        

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}