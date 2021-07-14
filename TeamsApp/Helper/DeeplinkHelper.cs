using Sinequa.Microsoft.Teams.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Sinequa.Microsoft.Teams.Helper
{
    public static class DeeplinkHelper
    {
    
        public static string GetTaskDeepLink(string strAppID,string baseUrl )
        {
            return string.Format("https://teams.microsoft.com/l/task/{0}?url={1}&height={2}&width={3}&title={4}&completionBotId={5}",
              strAppID,
              HttpUtility.UrlEncode(baseUrl + "/preview"),
              700,
              500,
              HttpUtility.UrlEncode("Insight Portal"),
              strAppID);
        }

        public static string GetPopUpDocCacheDeepLink(string strAppID, string baseUrl, string docCacheUrl)
        {
            UISettings PREVIEWSETTINGS = TaskModuleUIConstants.Preview;
            return string.Format("https://teams.microsoft.com/l/task/{0}?url={1}&height={2}&width={3}&title={4}&completionBotId={5}",
              strAppID,
              HttpUtility.UrlEncode($"{baseUrl}/preview?url={Convert.ToBase64String(Encoding.UTF8.GetBytes(docCacheUrl))}"),
              PREVIEWSETTINGS.Height,
              PREVIEWSETTINGS.Width,
              HttpUtility.UrlEncode(PREVIEWSETTINGS.Title),
              strAppID);
        }
    }
}
