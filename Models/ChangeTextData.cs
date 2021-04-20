using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ChangeText.WebApi.Models
{
    public class ChangeTextData
    {
        /// <summary>
        /// 文档路径
        /// </summary>
        public string DocUrl { get; set; }
        /// <summary>
        /// 绑定方案
        /// </summary>
        public int BindingMethod { get; set; }
        /// <summary>
        /// 行索引
        /// </summary>
        public int Line { get; set; }
        /// <summary>
        /// 列索引
        /// </summary>
        public int Column { get; set; }
        /// <summary>
        /// 内容
        /// </summary>
        public string Value { get; set; }
        /// <summary>
        /// 方位
        /// </summary>
        public string Location { get; set; }
        /// <summary>
        /// 搜索字段
        /// </summary>
        public string PdfField { get; set; }

        public static ChangeTextData FromJsonString(string jsonString)
        {
            ChangeTextData info = JsonConvert.DeserializeObject<ChangeTextData>(jsonString);
            return info;
        }


        public static ChangeTextData FromJson(JObject from)
        {
            JsonSerializer serializer = new JsonSerializer();
            ChangeTextData p = (ChangeTextData)serializer.Deserialize(new JTokenReader(from), typeof(ChangeTextData));
            return p;
        }


        public override string ToString()
        {
            string output = JsonConvert.SerializeObject(this);
            return output;
        }

    }

    public class ChangTestList
    {
        /// <summary>
        /// json数组
        /// </summary>
        public string CtdListString { get; set; }
    }
}