using System;
using System.Net.Http;
using System.Web.Http;
using Microsoft.Office.Interop.Word;
using ChangeText.WebApi.Models;

namespace ChangeText.WebApi.Controllers
{

    public class WordToPdfController : ApiController
    {
		
        private ApiTools tool = new ApiTools();

		[HttpPost]
		public HttpResponseMessage WToP(WordToPdfData wtp)
		{
            try
            {
                //建立 word application instance
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                //打开 word 文档
                Document wordDocument = appWord.Documents.Open(wtp.Sourcedocx);
                //转 PDF
                wordDocument.ExportAsFixedFormat(wtp.Targetpdf, WdExportFormat.wdExportFormatPDF);
                //关闭 word 文档
                wordDocument.Close();
                appWord.Quit();
            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "装换失败", ex.ToString());
            }

            return tool.MsgFormat(ResponseCode.成功, "装换成功", "0");
        }
	}

}