using System;
using System.Linq;
using ChangeText.WebApi.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChangeText.WebApi.Controllers
{
    public class WriteWordTableHandle
    {
        public void ChangeTextInCell(ChangeTextData ctData)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(ctData.DocUrl, true))
            {

                Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                TableRow row = table.Elements<TableRow>().ElementAt(Convert.ToInt32(ctData.Line));

                TableCell cell = row.Elements<TableCell>().ElementAt(Convert.ToInt32(ctData.Column));

                Paragraph p = cell.Elements<Paragraph>().First();

                Run r = p.Elements<Run>().First();

                Text t = r.Elements<Text>().First();

                t.Text = ctData.Value;
            }
        }

        public void ChangeTextInCell(string docUrl, string line, string column, string value)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(docUrl, true))
            {

                Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                TableRow row = table.Elements<TableRow>().ElementAt(Convert.ToInt32(line));

                TableCell cell = row.Elements<TableCell>().ElementAt(Convert.ToInt32(column));

                Paragraph p = cell.Elements<Paragraph>().First();

                Run r = p.Elements<Run>().First();

                Text t = r.Elements<Text>().First();
                t.Text = value;
            }
        }

    }
}
