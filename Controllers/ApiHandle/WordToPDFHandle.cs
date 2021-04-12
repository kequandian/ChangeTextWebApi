using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ChangeText.WebApi.Models;
using Microsoft.Office.Interop.Word;

namespace ChangeText.WebApi.Controllers.ApiHandle
{
    public class WordToPDFHandle
    {
        public void WToPdf(WordToPdfData wtp)
        {
            ////建立 word application instance
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            ////打开 word 文档
            var wordDocument = appWord.Documents.Open(wtp.Sourcedocx);
            ////转 PDF, 会自动创建同名的PDF格式文件
            wordDocument.ExportAsFixedFormat(wtp.Targetpdf, WdExportFormat.wdExportFormatPDF);
            ////关闭 word 文档
            wordDocument.Close();
            appWord.Quit();
        }
    }
}