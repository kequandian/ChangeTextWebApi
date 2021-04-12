using System;
using System.Net.Http;
using System.Web.Http;
using ChangeText.WebApi.Controllers.ApiHandle;
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
                WordToPDFHandle wordToPdfDataHandle = new WordToPDFHandle();
                wordToPdfDataHandle.WToPdf(wtp);
            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "转换失败", ex.ToString());
            }

            return tool.MsgFormat(ResponseCode.成功, "转换成功", "0");
        }
	}

}