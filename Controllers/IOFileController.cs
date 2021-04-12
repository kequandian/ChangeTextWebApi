using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace ChangeText.WebApi.Controllers
{
    public class IOFileController : ApiController
    {
        private ApiTools tool = new ApiTools();
        [HttpGet]
        public HttpResponseMessage DownloadFile(string fn)
        {
            string fileName = fn;
            try
            {
                if(fileName.LastIndexOf(".") == -1)
                {
                    fileName = string.Format("{0}.pdf", fn);
                }
                
                //string filePath = HttpContext.Current.Server.MapPath("~/") + "FileRoot\\" + "ReportTemplate.xlsx";
                
                string filePath = "D:\\workspace2015\\docxFile\\" + fileName;//测试下载的路劲

                FileStream stream = new FileStream(filePath, FileMode.Open);
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new StreamContent(stream);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = HttpUtility.UrlEncode(fileName)
                };
                response.Headers.Add("Access-Control-Expose-Headers", "FileName");
                response.Headers.Add("FileName", HttpUtility.UrlEncode(fileName));
                return response;

            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "下载失败", ex.ToString());
            }

            //return tool.MsgFormat(ResponseCode.成功, "下载成功", "0");
        }
    }
}