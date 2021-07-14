using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sinequa.Microsoft.Teams.Models
{
    public class PreviewTask
    {
        [JsonProperty("actionId")]
        public string ActionId { get; set; } = "preview";

        [JsonProperty("url")]
        public string Url 
        { 
            get; 
            set; 
        }
    }
}
