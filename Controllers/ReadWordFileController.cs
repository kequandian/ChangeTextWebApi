using ChangeText.WebApi.Models;
using docx.openxml;
using docx.openxml.models;
using System;
using System.Net.Http;
using System.Web.Http;

namespace ChangeText.WebApi.Controllers
{

    public class ReadWordFileController : ApiController
    {
		
        private ApiTools tool = new ApiTools();
		[HttpPost]
		public HttpResponseMessage ReadFile(FilePathInfo fp)
		{

            try
            {
                ReadWordFileHandle readWordFile = new ReadWordFileHandle();
                readWordFile.ReadWordTableToSave(fp.FilePath);

            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "读取失败", ex.ToString());
            }

            return tool.MsgFormat(ResponseCode.成功, "读取成功", "0");
        }
	}

}