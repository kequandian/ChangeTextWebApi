using System;
using System.Net.Http;
using System.Web.Http;
using ChangeText.WebApi.Controllers.ApiHandle;
using ChangeText.WebApi.Models;

namespace ChangeText.WebApi.Controllers
{

    public class ConvertToPdfController : ApiController
    {
		
        private ApiTools tool = new ApiTools();


        [HttpPost]
        public HttpResponseMessage ToPdf(HtmlToPdfData wth)
        {
            try
            {
                OfficeConvertHandle officeConvertHandle = new OfficeConvertHandle();
                officeConvertHandle.HtmlToPdf(wth);
            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "转换pdf失败", ex.ToString());
            }

            return tool.MsgFormat(ResponseCode.成功, "转换pdf成功", "0");
        }

    }

}