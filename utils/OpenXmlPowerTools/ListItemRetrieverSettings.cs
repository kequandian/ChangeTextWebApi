using ChangeText.WebApi.utils.OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ChangeText.WebApi.utils.OpenXmlPowerTools
{
    public class ListItemRetrieverSettings
    {
        public static Dictionary<string, Func<string, int, string, string>> DefaultListItemTextImplementations =
            new Dictionary<string, Func<string, int, string, string>>()
            {
                {"fr-FR", ListItemTextGetter_fr_FR.GetListItemText},
                {"tr-TR", ListItemTextGetter_tr_TR.GetListItemText},
                {"ru-RU", ListItemTextGetter_ru_RU.GetListItemText},
                {"sv-SE", ListItemTextGetter_sv_SE.GetListItemText},
                {"zh-CN", ListItemTextGetter_zh_CN.GetListItemText},
            };
        public Dictionary<string, Func<string, int, string, string>> ListItemTextImplementations;
        public ListItemRetrieverSettings()
        {
            ListItemTextImplementations = DefaultListItemTextImplementations;
        }
    }
}