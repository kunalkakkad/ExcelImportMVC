using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace ExcelImportDemo.Helpers
{
    class ExcelReaderDbManager
    {
        #region DataTable to Json

        public static dynamic ConvertDataTableToJSon(DataTable dt)
        {
            if (dt == null || dt.Rows.Count <= 0)
                return string.Empty;

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            List<Dictionary<string, object>> lstRows = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;
            foreach (DataRow dr in dt.Rows)
            {
                row = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    row.Add(col.ColumnName, dr[col]);
                }
                lstRows.Add(row);
            }
            return serializer.Serialize(lstRows);
        }

        public static dynamic ConvertJsonToDataTable(string strJson)
        {
            try
            {
                var result = JsonConvert.DeserializeObject(strJson, typeof(DataTable));

                return result;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        //Below code is for opening Excel file using DocumentFormat.xml.

        public static DataTable READExcel(string path, out int columnCount)
        {
            columnCount = 0;
            DataTable dataTable = new DataTable();
            try
            {
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(path, false))
                {
                    WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                    IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                    string relationshipId = sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Descendants<Row>();
                    foreach (Cell cell in rows.ElementAt(0))
                    {
                        dataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                    }

                    for (int i = 1; i < rows.Count(); i++)
                    {
                        DataRow dataRow = dataTable.NewRow();

                        int cellCounter = 0;
                        foreach (Cell cell in rows.ElementAt(i))
                        {
                            dataRow[cellCounter] = GetCellValue(spreadSheetDocument, cell);
                            cellCounter++;
                        }

                        dataTable.Rows.Add(dataRow);
                    }

                }
                //dataTable.Rows.RemoveAt(0);
                columnCount = dataTable.Columns.Count;
                return dataTable;


            }
            catch (Exception ex)
            {
                //DatabaseLog(ex);
                return dataTable;
            }
        }

        public static string GetCellValue(SpreadsheetDocument document, DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        {
            if (cell.CellValue == null)
                return string.Empty;

            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else if (cell.StyleIndex != null && (cell.StyleIndex.Value == (uint)1)) //Excel stores date value in Double data type and its cell's data type will always be null. so check if styleIndex = 1 then its date
            {
                double dValue = Convert.ToDouble(value);
                DateTime dtDate = DateTime.FromOADate(dValue);
                return dtDate.ToString("yyyy-MM-dd");
            }
            else
            {
                return value;
            }
        }
    }
}
