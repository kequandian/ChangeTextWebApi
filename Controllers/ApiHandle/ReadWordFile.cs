using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using docx.openxml.models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace docx.openxml
{
    public class ReadWordFile
    {
        public void ReadWordTableToSave(string filePath)
        {

            using (var document = WordprocessingDocument.Open(filePath, true))
            {
                var docPart = document.MainDocumentPart;

                var doc = docPart.Document;

                DocumentFormat.OpenXml.Wordprocessing.Table myTable = doc.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();

                List<List<string>> totalRows = new List<List<string>>();
                int rows = 0;
                int columns = 0;
                List<PositionInfo> pInfoList = new List<PositionInfo>() ;
                PositionInfo pInfo;

                foreach (TableRow row in myTable.Elements<TableRow>())
                {
                    rows++;
                    columns = 0;
                    List<string> tempRowValues = new List<string>();
                    foreach (TableCell cell in row.Elements<TableCell>())
                    {
                        columns++;
                        pInfo = new PositionInfo();
                        pInfo.RowIndex = rows;
                        pInfo.ColumnIndex = columns;
                        pInfo.Content = cell.InnerText;
                        pInfoList.Add(pInfo);
                    }
                }

            }

        }

    }
}
