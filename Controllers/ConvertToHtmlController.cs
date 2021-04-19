using System;
using System.Net.Http;
using System.Web.Http;
using ChangeText.WebApi.Controllers.ApiHandle;
using ChangeText.WebApi.Models;

namespace ChangeText.WebApi.Controllers
{

    public class ConvertToHtmlController : ApiController
    {
		
        private ApiTools tool = new ApiTools();


        [HttpPost]
        public HttpResponseMessage ToHtml(WordToHtmlData wth)
        {
            try
            {
                OfficeConvertHandle officeConvertHandle = new OfficeConvertHandle();
                officeConvertHandle.ConvertToHtml(wth);
            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "转换html失败", ex.ToString());
            }

            return tool.MsgFormat(ResponseCode.成功, "转换html成功", "0");
        }

    }

}