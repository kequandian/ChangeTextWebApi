using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using ChangeText.WebApi.Controllers.ApiHandle;
using ChangeText.WebApi.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChangeText.WebApi.Controllers
{

    public class DownloadFileController : ApiController
    {
		
        private ApiTools tool = new ApiTools();

        [HttpPost]
        public HttpResponseMessage Download(ChangTestList data)
        {
            string docUrl = "D:\\workspace2015\\docxFile\\对经营高危险性体育项目的行政许可.docx";
            try
            {

                WriteWordTableHandle writeWordTable = new WriteWordTableHandle();
                ChangeTextData ctData = null;

                #region 处理json数组
                JArray info = (JArray)JsonConvert.DeserializeObject(data.CtdListString);

                if (info != null)
                {
                    for (int i = 0; i < info.Count; i++)
                    {
                        JObject joRet = (JObject)info[i];
                        if (joRet["bindingMethod"].ToString().Equals("1"))
                        {
                            //ctData = ChangeTextData.FromJson(joRet);
                            //ctData.DocUrl = docUrl;
                            ctData = new ChangeTextData()
                            {
                                DocUrl = docUrl,
                                Line = joRet["line"].ToString() != string.Empty ? Convert.ToInt32(joRet["line"]) : -1,
                                Column = joRet["column"].ToString() != string.Empty ? Convert.ToInt32(joRet["column"]) : -1,
                                Value = joRet["value"].ToString() != string.Empty ? joRet["column"].ToString() : "",

                            };
                            if (ctData.Line >= 0 && ctData.Column >= 0)
                            {
                                writeWordTable.ChangeTextInCell(ctData);
                            }
                        }
                    }
                }
                #endregion

                #region word 转 html
                OfficeConvertHandle officeConvertHandle = new OfficeConvertHandle();
                WordToHtmlData wth = new WordToHtmlData()
                {
                    Sourcedocx = docUrl,
                    TargetDirectory = "D:\\workspace2015\\docxFile"
                };
                officeConvertHandle.ConvertToHtml(wth);
                #endregion

                #region html 转 pdf
                HtmlToPdfData htp = new HtmlToPdfData() { 
                    SourceHtml = "D:\\workspace2015\\docxFile\\对经营高危险性体育项目的行政许可.html",
                    TargetPdf = "D:\\workspace2015\\docxFile\\对经营高危险性体育项目的行政许可.pdf"
                };
                officeConvertHandle.HtmlToPdf(htp);
                #endregion

                #region 下载文档
                string fn = "对经营高危险性体育项目的行政许可.pdf";

                string fileName = fn
                    ;
                if (fileName.LastIndexOf(".") == -1)
                {
                    fileName = string.Format("{0}.pdf", fn);
                }


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
                #endregion

            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "操作失败", ex.ToString());
            }

            //return tool.MsgFormat(ResponseCode.成功, "操作成功", "0");
        }

    }

}