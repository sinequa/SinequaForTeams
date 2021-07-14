using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using System;

namespace Sinequa.Microsoft.Teams.Pages
{
    public class PreviewModel : PageModel
    {
        

        public string PreviewUrl { get; private set; }

        public void OnGet(string url) {
            PreviewUrl = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(url));
        }
    }
}
