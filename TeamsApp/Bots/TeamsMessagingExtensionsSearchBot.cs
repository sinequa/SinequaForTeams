// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Newtonsoft.Json.Linq;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.IO;
using AdaptiveCards.Templating;
using System;
using System.Text;
using System.Net;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Activity = Microsoft.Bot.Schema.Activity;
using Sinequa.Microsoft.Teams.Helper;
using Sinequa.Microsoft.Teams.Models;
using Microsoft.AspNetCore.WebUtilities;
using Sinequa.Search;
using Sinequa.Common;
using Sinequa.Configuration;
using Sinequa.Configuration.SBA;
using Sinequa.Engine.Client;
using Sinequa.Search.JsonMethods;
using Sinequa.Plugins;
using TeamsMessagingExtensionsSearch.SinequaPlugin;
using System.Reflection;
using Microsoft.Bot.Connector;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.BotBuilderSamples;
using Microsoft.Bot.Connector.Authentication;

namespace Sinequa.Microsoft.Teams.Bots
{
    public class TeamsMessagingExtensionsSearchBot<T> : DialogBot<T> where T:Dialog
    {

        private readonly ILogger _logger;
        private readonly string _baseUrl;
        private readonly string host;
        private readonly int port;
        private readonly string sinequa_app_name;
        private readonly string sinequa_ws_query_name;
        private readonly string _microsoftAppId;
        private readonly string domainName;
        private readonly string jwtBearerToken;
        private readonly string SINEQUA_LOGO = "https://ga1.imgix.net/logo/o/155731-1593616014-6035028?ixlib=rb-1.0.0&ch=Width%2CDPR&auto=format";
        private readonly int pageSize = 5;
        private readonly string _connectionName;



        public TeamsMessagingExtensionsSearchBot(ConversationState conversationState, UserState userState, T dialog, IConfiguration configuration, ILogger<TeamsMessagingExtensionsSearchBot<T>> logger)
            :base(conversationState, userState, dialog, logger) 
        {
            
            _logger = logger;

            if (_logger != null)
            {
                _logger.LogDebug("TeamsMessagingExtensionsSearchBot ctor");
            }

            //Read the configuation values from appSettings.json
            host = configuration["Sinequa:HostName"];
            if (!int.TryParse(configuration["Sinequa:CustomPort"], out port))
            {
                port = 443;
            }

            sinequa_app_name = configuration["Sinequa:AppName"];
            sinequa_ws_query_name = configuration["Sinequa:WSQueryName"];
            domainName = configuration["Sinequa:Domain"];
            jwtBearerToken = configuration["JWT"];// _configuration["Sinequa:JWTAccessToken"];
            _baseUrl = configuration["Sinequa:BaseUrl"];
            _connectionName = configuration["Sinequa:ConnectionName"];

            if (string.IsNullOrEmpty(jwtBearerToken))
            {
                _logger.LogError("Cannot retrieve JWT Bearer Token");
            }
            
            _microsoftAppId = configuration["MicrosoftAppId"];
            string MicrosoftAppPassword = configuration["MicrosoftAppPassword"];

            if (_logger != null)
            {
                _logger.LogInformation("Sinequa:HostName " + host);
                _logger.LogInformation("Sinequa:CustomPort " + port);
                _logger.LogInformation("Sinequa:AppName " + sinequa_app_name);
                _logger.LogInformation("Sinequa:WSQueryName " + sinequa_ws_query_name);
                _logger.LogInformation("Sinequa:Domain " + domainName);
                _logger.LogInformation("Sinequa:BaseUrl " + _baseUrl);

                
                _logger.LogInformation("MicrosoftAppId " + _microsoftAppId);
                //_logger.LogInformation("MicrosoftAppPassword " + MicrosoftAppPassword.Substring(0, 5) + "[...]");
                //_logger.LogInformation("jwtBearerToken " + jwtBearerToken.Substring(0, 5) + "[...]");
                //
            }


        }



        // If there is no
        // nail associated with a document, then use one of the below icons ( based on the document type) to display on the Card.
        static IDictionary<string, string> _mappings = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase)
        {
            { "docx", "https://img.icons8.com/color/2x/microsoft-word-2019.png" },
            { "doc", "https://img.icons8.com/color/2x/microsoft-word-2019.png" },
            { "ppt", "https://img.icons8.com/color/2x/ms-powerpoint.png" },
            { "pptx", "https://img.icons8.com/color/2x/ms-powerpoint.png" },
            { "xlsx", "https://img.icons8.com/color/2x/microsoft-excel-2019.png" },
            { "xls", "  https://img.icons8.com/color/2x/microsoft-excel-2019.png"},
            { "pdf", "https://img.icons8.com/color/2x/pdf.png"},
            { "htm","https://img.icons8.com/fluent/2x/domain.png" },
            { "html","https://img.icons8.com/fluent/2x/domain.png" },
            { "xml","https://img.icons8.com/dusk/2x/xml-file.png" }

        };
        private static IDictionary<string, string> EXT_TO_ICON = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase)
        {
            { "docx", "doc32flat.png" },
            { "dot", "doc32flat.png" },
            { "doc", "doc32flat.png" },
            { "rtf", "doc32flat.png" },
            { "ppt", "ppt32flat.png" },
            { "pptx", "ppt32flat.png" },
            { "pptm", "ppt32flat.png" },
            { "xlsx", "xls32flat.png" },
            { "xls", "xls32flat.png"},
            { "pdf", "pdf32flat.png"},
            { "htm","html32flat.png" },
            { "html","html32flat.png" },
            { "xml","xml32flat.png" },
            { "txt","txt32flat.png" }
        };
        private static string DEFAULT_ICON = "any32flat.png";


        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            IActivity iActvity = turnContext.Activity;

            string contentString = string.Empty;
            switch (action.CommandId)
            {
                case "searchQuery":
                    //extract query content from data
                    JObject obj = action?.Data as JObject;
                    contentString = obj?.GetValue("searchQuery")?.Value<String>();
                    break;
                //case "searchForIt":
                //    //extract content from payload
                //    string content = action?.MessagePayload?.Body?.Content;
                //    string ct = action?.MessagePayload?.Body?.ContentType;
                //    contentString = ParseContent(content, ct);
                //    break;
                default:
                    throw new NotImplementedException($"Invalid CommandId: {action.CommandId}");
            }


            var queryText = contentString;
            var user = turnContext?.Activity.From;

            var blob_storage_thumbnailurl = SINEQUA_LOGO + "&h=50&w=50";

            _logger.LogInformation("Teams UserID  = " + user.Id + "Teams Username :user.Name = " + user.Name + " Azure AD ID= " + user.AadObjectId);

            //var packages = await FindResults(user.Id, queryText);
            var records = FindResults(user.AadObjectId, queryText);

            var attachments = records.Select(record =>
            {
                var previewCard = new HeroCard { Title = record.Item7, Text = record.Item3, Tap = new CardAction { Type = "invoke", Value = record } };
                if (!string.IsNullOrEmpty(record.Item10) && _mappings.ContainsKey(record.Item10))
                {
                    _mappings.TryGetValue(record.Item10, out blob_storage_thumbnailurl);
                }
                previewCard.Images = new List<CardImage>() { new CardImage(blob_storage_thumbnailurl, "Icon", "OpenUrl") };


                var attachment = new MessagingExtensionAttachment
                {
                    ContentType = HeroCard.ContentType,
                    Content = new HeroCard { Title = record.Item7, Text = record.Item3 },
                    Preview = previewCard.ToAttachment()
                };
                return attachment;
            }).ToList();

            return new MessagingExtensionActionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = attachments,
                }
            };


            //return await Task.FromResult(new MessagingExtensionActionResponse());
        }

        protected override async Task OnSignInInvokeAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            await _dialog.RunAsync(turnContext, _conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
        }
        protected override async Task OnTokenResponseEventAsync(ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
        {
            await _dialog.RunAsync(turnContext, _conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
        }

        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
        {
            await base.OnTurnAsync(turnContext, cancellationToken);

            // Save any state changes that might have occurred during the turn.
            await _conversationState.SaveChangesAsync(turnContext, false, cancellationToken);
            await _userState.SaveChangesAsync(turnContext, false, cancellationToken);
        }

        //Called when user is interacting with the chat in the bot channel 
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {

            if (turnContext.Activity.Value != null) // TODO Message from a card.
            {
                var obj = (JObject)turnContext.Activity.Value;
                var answer = obj["Answer"]?.ToString();
                //var choices = obj["Choices"]?.ToString();
                //await turnContext.SendActivityAsync(MessageFactory.Text($"{turnContext.Activity.From.Name} answered '{answer}' and chose '{choices}'."), cancellationToken);
            }
            else //regular text message
            {
                var queryText = turnContext?.Activity?.Text;

                if (String.IsNullOrWhiteSpace(queryText))
                {
                    if (turnContext.Activity?.Attachments?.Count >= 1)
                    {
                        await turnContext.SendActivityAsync("Sorry, I can only process queries from simple text. I can't handle attachements such as cards :(");
                    }
                    else
                    {
                        await turnContext.SendActivityAsync("Please enter a valid search query (Empty search not availaible from Teams Sinequa App)");
                    }
                    return;
                }

                var user = turnContext?.Activity.From;
                var blob_storage_thumbnailurl = SINEQUA_LOGO + "&h=50&w=50";

                _logger.LogInformation("Teams UserID  = " + user.Id + "Teams Username :user.Name = " + user.Name + " Azure AD ID= " + user.AadObjectId);

                //var packages = await FindResults(user.Id, queryText);
                var records = FindResults(user.AadObjectId, queryText);

                //var records = new List<(string, string, string, string, string, string, string, string, string, string, string)>();

                //Response from FindResults- records=> id, authors, smallsummaryhtml, modified, treepath, url1, title, thumbnail, objectType, fileext,queryText

                // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
                // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.
                if(records == null)
                {
                    await turnContext.SendActivityAsync("Search error (is engine started ?)");
                    return;
                }


                var attachments = records.Select(record =>
                {
                    var template = GetAdaptiveCardTemplate("adaptive2.json");
                    var docCacheUrl = GetDocCacheURL(this.host, this.port, this.sinequa_app_name, this.sinequa_ws_query_name, record.Item1, queryText)?.AbsoluteUri;
                    var docDirectLinkUrl = GetDirectLinkURL(this.host, this.port, this.sinequa_app_name, this.sinequa_ws_query_name, record.Item1, queryText)?.AbsoluteUri;
                    DateTime modifiedDT = DateTime.Now;
                    if (!DateTime.TryParse(record.Item4, out modifiedDT))
                    {
                        modifiedDT = DateTime.Now;
                    }
                    Attachment image = ImageFromExtension(record.Item10);
                    var myData = new
                    {
                        DocTitle = record.Item7,
                        Summary = StripHtml(record.Item3), // ConvertHTMLToMarkdown(relevantExtracts),  //<b>...</b>  //**...**  //encoding issue on author
                        ThumbnailUrl = image.ContentUrl,
                        PreviewUrl = docCacheUrl,
                        DirectlinkUrl = docDirectLinkUrl,// record.Item6,
                        SourceTreepath = SourceFromTreePath(record.Item5),
                        AuthorName = StringCleanup(record.Item2),  //encoding issue 
                        FileType = record.Item9,
                        Modified = modifiedDT.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ssK")
                    };
                    var cardJson = template.Expand(myData);
                    var adaptiveCardAttachment = new Attachment()
                    {
                        ContentType = "application/vnd.microsoft.card.adaptive",
                        Content = JsonConvert.DeserializeObject(cardJson),
                    };

                    return adaptiveCardAttachment;
                }).ToList();

                
                var activity = MessageFactory.Carousel(attachments, attachments?.Count > 0 ? $"Here's what I found\r\n" : "Your search did not match any documents");
                await turnContext.SendActivityAsync(MessageFactory.Carousel(attachments, attachments?.Count > 0 ? $"Here's what I found\r\n" : "Your search did not match any documents"), cancellationToken);

            }

        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var welcomeText = "Hello and welcome!";
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText, welcomeText), cancellationToken);
                }
            }
        }



        /** 
         * MAIN SEARCH ENTRY POINT @SNQA ....)
         * This method is invoked when the user enters the query string in the Sinequa For Teams App. It triggers a Sinequa Backend Query , gets the responses
            and displays the search result list , where each row of the result is wrapped in a Hero card, that are wrapped in MessagingExtensionAttachment objects.
            A 'Tap' Action on one of the results will invoke the OnTeamsMessagingExtensionSelectItemAsync , to display an Adaptive Card for the selected result,
            that can in turn be copy /pasted to a chat.
         **/
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
        {
            var queryText = query?.Parameters?[0]?.Value as string ?? string.Empty;
            var user = turnContext?.Activity.From;

            _logger.LogDebug("Teams UserID  = " + user.Id + "Teams Username :user.Name = " + user.Name + " Azure AD ID= " + user.AadObjectId);

            //var packages = await FindResults(user.Id, queryText);
            var records = FindResults(user.AadObjectId, queryText);

            //Response from FindResults- packages=> id, authors, smallsummaryhtml, modified, treepath, url1, title, thumbnail, objectType, fileext,queryText

            // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
            // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.
            var attachments = records.Select(record =>
            {
                DateTime modifiedDT = DateTime.Now;
                if (!DateTime.TryParse(record.Item4, out modifiedDT))
                {
                    modifiedDT = DateTime.Now;
                }
                string strModified = modifiedDT.ToString("ddd, MMM d yyyy");


                //var previewCard = new HeroCard { Title = package.Item7, Text = package.Item3, Tap = new CardAction { Type = "invoke", Value = package } };
                var previewCard = new ThumbnailCard
                {
                    Title = record.Item7,
                    Subtitle = $"{StringCleanup(record.Item2)} - {strModified}",
                    Text = StripHtml(record.Item3, false),
                    Tap = new CardAction { Type = "invoke", Value = record }
                };

                Attachment image = ImageFromExtension(record.Item10);
                previewCard.Images = new List<CardImage>() { new CardImage(image.ContentUrl, record.Item6, "OpenUrl") };


                var attachment = new MessagingExtensionAttachment
                {
                    ContentType = HeroCard.ContentType,
                    Content = new HeroCard { Title = record.Item7, Text = record.Item3 },
                    Preview = previewCard.ToAttachment()
                };
                return attachment;
            }).ToList();


            #region need reengineering
            // Creating Deep Link URLs to provide Links to the Personal Tabs, if the user wants to run a more Advanced Search.
            //TODO: Revisit this whole part. This can't remain hardcoded - Externalize this as an adaptaive card resource...
            // edafa5ca-ded7-4b90-9735-4e7737bd841b refers to the old app id ...
            //var medTechDeepLinkURL = "https://teams.microsoft.com/l/entity/edafa5ca-ded7-4b90-9735-4e7737bd841b/sinequaMedTech?label=medTech";
            //var covidSearchDeepLinkURL = "https://teams.microsoft.com/l/entity/edafa5ca-ded7-4b90-9735-4e7737bd841b/covidSearchDev?label=covidSearch";

            //var insightDeepLinkURL = DeeplinkHelper.GetTaskDeepLink(_microsoftAppId, _baseUrl);
            //var lastPreviewCard = new HeroCard
            //{
            //    Title = "<p style='color:darkpurple;font-size:14px'><b> Advanced Search - Click Here to Search Within the Tabs!</b></p>",
            //    Buttons = new List<CardAction>
            //        {
            //            new CardAction { Type = ActionTypes.OpenUrl, Title = "<p style='color:darkpurple;font-size:12px'> <b>Insight Portal</p>", Image="https://beezysinequa.blob.core.windows.net/demo-thumbnails/formats/covid-light-300.png",Value = insightDeepLinkURL }
            //        }
            //};
            //var lastAttachment = new MessagingExtensionAttachment
            //{
            //    ContentType = HeroCard.ContentType,
            //    Content = lastPreviewCard,
            //    Preview = lastPreviewCard.ToAttachment(),

            //};
            //attachments.Insert(0, lastAttachment);
            #endregion

            // The list of MessagingExtensionAttachments must we wrapped in a MessagingExtensionResult wrapped in a MessagingExtensionResponse.
            return new MessagingExtensionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = attachments,
                }
            };
        }

        private static Attachment ImageFromExtension(string extension)
        {
            String imageName;
            if (!EXT_TO_ICON.TryGetValue(extension, out imageName))
            {
                imageName = DEFAULT_ICON;
            }
            Attachment image = Resource2InlineAttachment(imageName);
            return image;
        }

        private static Uri GetDocCacheURL(string host, int port, string appName, string queryWSName, string id, string queryText)
        {
            var urlEncodedID = System.Net.WebUtility.UrlEncode(id);
            var urlEncodedJSonPayload = System.Net.WebUtility.UrlEncode($"{{\"name\":{ JsonConvert.SerializeObject(queryWSName)},\"text\":{JsonConvert.SerializeObject(queryText)}}}");
            UriBuilder ub = new UriBuilder();
            ub.Scheme = "https";
            ub.Host = host;
            ub.Port = port;
            ub.Path = $"/app/{appName}/";
            //workaround for the sharp sign (otherwise urlencoded by uribuilder), tried the fragment part  w/o success...
            ub.Query = $"#/preview?id={urlEncodedID}&query={urlEncodedJSonPayload}";

            //ub.Path = $"/xdownload/html/~" + System.Net.WebUtility.UrlEncode(System.Net.WebUtility.UrlEncode("{\"query\":{ \"name\":\"" + queryWSName + "\",\"text\":\"" + queryText + "\"},\"app\":\"" + appName + "\",\"id\":\"" + id + "\"}"))+"~/file.htm";
            return ub.Uri;
        }

        private static Uri GetDirectLinkURL(string host, int port, string appName, string queryWSName, string id, string queryText)
        {
            var urlEncodedID = System.Net.WebUtility.UrlEncode(id);
            var urlEncodedJSonPayload = System.Net.WebUtility.UrlEncode($"{{\"name\":{ JsonConvert.SerializeObject(queryWSName)},\"text\":{JsonConvert.SerializeObject(queryText)}}}");
            UriBuilder ub = new UriBuilder();
            ub.Scheme = "https";
            ub.Host = host;
            ub.Port = port;
            ub.Path = $"/app/{appName}/";
            //workaround for the sharp sign (otherwise urlencoded by uribuilder), tried the fragment part  w/o success...
            ub.Query = $"#/preview?id={urlEncodedID}&query={urlEncodedJSonPayload}";
            return ub.Uri;
        }

        //display the adaptive card  
        protected override Task<MessagingExtensionResponse> OnTeamsMessagingExtensionSelectItemAsync(ITurnContext<IInvokeActivity> turnContext, JObject query, CancellationToken cancellationToken)
        {

            var user = turnContext?.Activity.From;

            // The Preview card's Tap should have a Value property assigned, this will be returned to the bot in this event. 
            var (id, authors, relevantExtracts, modified, treepath, url1, title, thumbnailUrl, objectType, fileext, text) = query.ToObject<(string, string, string, string, string, string, string, string, string, string, string)>();


            var docCacheURL = GetDocCacheURL(this.host, this.port, this.sinequa_app_name, this.sinequa_ws_query_name, id, text)?.AbsoluteUri;
            _logger.LogDebug("Doc Cache URL = " + docCacheURL.ToString());
            //Default Thumnail image - Sinequa Logo
            var blob_storage_thumbnailurl = SINEQUA_LOGO;

            if (!string.IsNullOrEmpty(fileext) && _mappings.ContainsKey(fileext))
            {
                _mappings.TryGetValue(fileext, out blob_storage_thumbnailurl);
            }
            string imgStr = null;
            if (!string.IsNullOrEmpty(thumbnailUrl))
            {
                _logger.LogDebug("thumbnailUrl = " + thumbnailUrl);
                var imageData = GetBase64Thumbnail(user.AadObjectId, thumbnailUrl);
                imgStr = imageData.Result;
            }
            //TODO Refactor  
            Attachment image;
            if (!string.IsNullOrEmpty(imgStr))
            {
                image = ImageData2InlineAttachment(imgStr);
            }
            else
            {
                image = ImageFromExtension(fileext);
            }
            //===========Nico:
            //Just replaced relevance by objectType that can drive the layout by matching with an adaptive card template.
            //=======================
            //var relevance = globalrelevance;
            //var result = float.TryParse(globalrelevance, out var relevancePercentage);
            //if (result)
            //    relevancePercentage = (float)Math.Round(relevancePercentage * 100, 2);
            //if (relevancePercentage > 100) //handles relevance trasnform scripts which alter the relevance to values that exceed 100...
            //{
            //    relevancePercentage = 100;
            //}
            if (treepath?.Length > 0)
                treepath = StringCleanup(treepath);

            if (authors?.Length > 0)
                authors = StringCleanup(authors);

            //Adaptive Card Logic
            //The Adaptive Card uses adaptiveCard.json as the template file. You must change the template to change the Adaptive Card design.

            AdaptiveCardTemplate template = GetAdaptiveCardTemplate($"{objectType?.ToLowerInvariant()}.json", "adaptiveCard.json");



            //Convert Summary content from HTML to Markdown as needed by Adaptive Cards
            var myData = new
            {
                DocTitle = title,
                Summary = StripHtml(relevantExtracts), // ConvertHTMLToMarkdown(relevantExtracts),  //<b>...</b>  //**...**  //encoding issue on author
                ThumbnailUrl = image.ContentUrl, //blob_storage_thumbnailurl,
                PreviewUrl = GetDirectLinkURL(this.host, this.port, this.sinequa_app_name, this.sinequa_ws_query_name, id, "")?.AbsoluteUri, //DeeplinkHelper.GetPopUpDocCacheDeepLink(this._microsoftAppId, this._baseUrl, docCacheURL), //docCacheURL,
                DirectlinkUrl = url1,
                SourceTreepath = treepath,
                AuthorName = StringCleanup(authors),  //encoding issue 
                FileType = fileext
        };

            var previewCard = new HeroCard { Title = title, Text = relevantExtracts };
            // "Expand" the template - this generates the final Adaptive Card payload
            var cardJson = template.Expand(myData);
            var adaptiveCardattachment = new MessagingExtensionAttachment
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(cardJson),
                Preview = previewCard.ToAttachment()

            };

            _logger.LogInformation("adaptiveCardattachment Content = " + adaptiveCardattachment.Content.ToString());

            return Task.FromResult(new MessagingExtensionResponse
            {
                ComposeExtension = new MessagingExtensionResult
                {
                    Type = "result",
                    AttachmentLayout = "list",
                    Attachments = new List<MessagingExtensionAttachment> { adaptiveCardattachment }
                }
            });
        }

        private static AdaptiveCardTemplate GetAdaptiveCardTemplate(string templateName, string defaultTemplate=null)
        {
            var fileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("TeamsMessagingExtensionsSearch.Resources."+templateName);

            string adaptiveCardJson;
            AdaptiveCardTemplate template;
            try
            {
                using (StreamReader reader = new StreamReader(fileStream))
                {
                    adaptiveCardJson = reader.ReadToEnd();
                }

                //string[] paths = { ".", "Resources", templateName };
                //if (defaultTemplate!=null && !File.Exists(Path.Combine(paths))) {
                //    paths = new string[] { ".", "Resources", defaultTemplate };
                //}
                //var adaptiveCardJson = File.ReadAllText(Path.Combine(paths));
                template = new AdaptiveCardTemplate(adaptiveCardJson);
            }catch(Exception ex)
            {
                if (defaultTemplate.Equals(templateName))
                    throw;
                return GetAdaptiveCardTemplate(defaultTemplate);
            }
            return template;
        }

        /** Link UnFurling  : If you paste a hyperlink that contains the domain *.sinequa.com, the link will be unfurled to display in a Thumbnail card.
         * Example : Try pasting a Document Cache link in the Teams Chat window.
         * The user can click on the card to open the link in a new tab.  This is helpful if you are selecting one of the search results in the Personal Tab and
         * pasting the Doc Cache Link in the chat message to share the doc with another user. 
         * The domains where link unfurling will apply are listed in the manifest.json in the "messageHandlers" section. Edit the list to add more domains as needed.
         * Refer to - https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/link-unfurling  for more details. 
         **/
        protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
        {
            if (query == null)
                throw new ArgumentNullException("query");
            if (turnContext == null)
                throw new ArgumentNullException("turnContext");

            var user = turnContext.Activity?.From;
            //TODO add safeguards about the url itself, not sure if Teams actually filters out abnormal URLs. 

            Uri uri = new Uri(query.Url);
            if (uri != null && !string.IsNullOrWhiteSpace(uri.AbsolutePath))
            {

                // [...]/app/insight/#/...
                if (!string.IsNullOrEmpty(uri.Fragment)
                    && uri.AbsolutePath.Equals($"/app/{this.sinequa_app_name}/", StringComparison.InvariantCultureIgnoreCase))
                {

                    ///#/preview?id=%2FI_Google%2FI_GoogleDrives%2F%7C1a6gOZfJcDiUQA4Rne2gsR9attfng5dRa&query=%7B%22name%22:%22insight_query%22,%22text%22:%22keywords%22%7D
                    // [...]/app/insight/#/search?query=....
                    //search query => returns carousel
                    if (uri.Fragment.StartsWith("#/search?query=", StringComparison.InvariantCultureIgnoreCase))
                    {

                        string decodedFragment = System.Web.HttpUtility.HtmlDecode(uri.Fragment);
                        string jsonPayload = decodedFragment.Substring("#/search?".Length);
                        jsonPayload = WebUtility.UrlDecode(jsonPayload);
                        if (jsonPayload.Length > 0)
                        {
                            JObject obj = JObject.Parse(jsonPayload);
                            JObject root = JObject.Parse($"{{\"app\":{JsonConvert.SerializeObject(this.sinequa_app_name)}}}");
                            root["query"] = obj;

                            var queryTextHint = obj?.GetValue("text")?.Value<String>();
                            //var records = await this.FindResultsFromPayload(user.AadObjectId, root.ToString(), queryTextHint);

                            //record =>
                            //id, authors, smallsummaryhtml, modified, treepath, url1, title, thumbnail, objectType, fileext,queryText
                            JObject rootObject = await this.FindResultsFromPayloadEx(user.AadObjectId, root.ToString(), queryTextHint);
                            var resCount = rootObject["totalRowCount"]?.ToObject<int>();
                            var record = rootObject["records"]?.First;
                            var docCacheUrl = GetDocCacheURL(this.host, this.port, this.sinequa_app_name, this.sinequa_ws_query_name, record["id"].ToString(), queryTextHint)?.AbsoluteUri;
                            DateTime modifiedDT = DateTime.Now;
                            if (!DateTime.TryParse(record["modified"].ToString(), out modifiedDT))
                            {
                                modifiedDT = DateTime.Now;
                            }

                            Attachment img = null;
                            string thumbnailUrl = record["thumbnailUrl"]?.ToString();
                            if (!string.IsNullOrEmpty(thumbnailUrl))
                            {
                                _logger.LogDebug("thumbnailUrl = " + thumbnailUrl);
                                var imageData = GetBase64Thumbnail(user.AadObjectId, thumbnailUrl);
                                img = ImageData2InlineAttachment(imageData.Result);
                            }
                            if (img == null)
                            {
                                img = ImageFromExtension(record["fileext"].ToString());
                            }

                            string strModified = modifiedDT.ToString("ddd, MMM d yyyy");
                            var heroCard2 = new ThumbnailCard
                            {
                                Title = $"Query - {queryTextHint} ",//record.Item7,
                                Subtitle = $"Doc#1 out of {resCount}: {record["title"]})",
                                Text = $"{StringCleanup(record["authors"]?.ToString())}\r\n{StripHtml(record["relevantExtracts"].ToString(),false)}",
                                Images = new List<CardImage> { 
                                            new CardImage(img.ContentUrl, "", null),
                                },
                                Tap = new CardAction { Type = "invoke", Value = null },
                                Buttons = new[] {
                                        new CardAction(ActionTypes.OpenUrl, "Open", null, "Original", "Original", record["url1"].ToString() ),
                                        new CardAction(ActionTypes.OpenUrl, "Preview", null, "Preview", "Preview", docCacheUrl),
                                        new CardAction(ActionTypes.OpenUrl, "Run query", null, "...", "...", uri.AbsoluteUri)
                                    }.ToList()
                            };

                            var attachments = new MessagingExtensionAttachment(ThumbnailCard.ContentType, null, heroCard2);
                            var msgngExtRslt = new MessagingExtensionResult("list", "result", new[] { attachments });

                            return await Task.FromResult(new MessagingExtensionResponse(msgngExtRslt));
                        }
                    }
                    else if (uri.Fragment.StartsWith("#/preview?", StringComparison.InvariantCultureIgnoreCase))
                    {
                        string decodedFragment = System.Web.HttpUtility.HtmlDecode(uri.Fragment);
                        var queryParameters = QueryHelpers.ParseQuery(decodedFragment.Substring("#/preview?".Length));
                        var docid = queryParameters?["id"];
                        if (!string.IsNullOrWhiteSpace(docid))
                        {
                            var queryDefJson = queryParameters?["query"];
                            JObject obj = JObject.Parse(queryDefJson);
                            var queryTextHint = obj?.GetValue("text")?.Value<String>();
                            string previewPayload = $"{{\"app\":{JsonConvert.SerializeObject(this.sinequa_app_name)},\"action\":\"get\",\"id\":{JsonConvert.SerializeObject(docid)},\"query\":{{ \"name\":{JsonConvert.SerializeObject(this.sinequa_ws_query_name)},\"text\":{JsonConvert.SerializeObject(queryTextHint)}}}}}";
                            Uri docCacheUrl = GetDocCacheURL(this.host, this.port, this.sinequa_app_name, this.sinequa_ws_query_name, docid, queryTextHint);


                            var root = await this.DoRequest(user.AadObjectId, previewPayload, "preview");
                            var record = root?["record"];

                            string thumbnailUrl = record["thumbnailUrl"]?.ToString();
                            Attachment imgThumbnail = null;
                            if (!string.IsNullOrEmpty(thumbnailUrl))
                            {
                                _logger.LogInformation("thumbnailUrl = " + thumbnailUrl);
                                var imageData = GetBase64Thumbnail(user.AadObjectId, thumbnailUrl);
                                imgThumbnail = ImageData2InlineAttachment(imageData.Result);
                            }
                            Attachment imgFileExt = ImageFromExtension(record["fileext"].ToString());

                            DateTime modifiedDT = DateTime.Now;
                            if (!DateTime.TryParse(record["modified"].ToString(), out modifiedDT))
                            {
                                modifiedDT = DateTime.Now;
                            }
                            string strModified = modifiedDT.ToString("ddd, MMM d yyyy");

                            var heroCard2 = new ThumbnailCard
                            {
                                Title = record["title"]?.ToString(),
                                Subtitle = $"{StringCleanup(record["authors"]?.ToString())} - {strModified}",
                                Images = new List<CardImage> { new CardImage(imgThumbnail != null ? imgThumbnail.ContentUrl : imgFileExt.ContentUrl, record["fileext"].ToString(), "OpenUrl") },
                                Tap = new CardAction { Type = "invoke", Value = record },
                                Buttons = new[] {
                                        new CardAction(ActionTypes.OpenUrl, "Original", null, "Original", "Original", record["url1"]?.ToString()),
                                        new CardAction(ActionTypes.OpenUrl, "Preview", null, "Preview", "Preview", DeeplinkHelper.GetPopUpDocCacheDeepLink(this._microsoftAppId, this._baseUrl, docCacheUrl.AbsoluteUri))
                                    }.ToList()
                            };
                            var attachments = new MessagingExtensionAttachment(HeroCard.ContentType, null, heroCard2);
                            var msgngExtRslt = new MessagingExtensionResult("list", "result", new[] { attachments });

                            return await Task.FromResult(new MessagingExtensionResponse(msgngExtRslt));
                        }
                    }
                    //TODO Handle more cases ?  
                    
                }

            }

            _logger.LogInformation("Link Unfurling - query = " + query + " query.url = " + query.Url);
            //Catch all !
            var heroCard = new HeroCard
            {
                Title = "Oops...",
                Text = $"We don't know how to handle this link  {query.Url}]",
                Images = new List<CardImage> { new CardImage(SINEQUA_LOGO, "Icon", "OpenUrl") },
                Tap = new CardAction { Type = "invoke", Value = query.Url }
            };

            var messagingExtensionAtt = new MessagingExtensionAttachment(HeroCard.ContentType, null, heroCard);
            var result = new MessagingExtensionResult("list", "result", new[] { messagingExtensionAtt });
            return await Task.FromResult(new MessagingExtensionResponse(result));
        }

        /**
         * Invokes a webservice query to Sinequa and Returns the results. The search call will be made using the Default Network Credentials of the user.
         * Works fine for the purposes of the demo. If the code is hosted on a different server, the logic will change to pass the user's credentials (may be use JWT) . 
         * */

        private IEnumerable<(string, string, string, string, string, string, string, string, string, string, string)> FindResults(string userid, string text)
        {

            CC cc = CC.Current;

            //CCPrincipal userPrincipal = cc.GetPrincipalAny(userid, domain);
            CCPrincipal userPrincipal = cc.GetPrincipalByUserIds(userid);




            //TODO : select a more specific engine to match the needs
            CCEngine ccEngine = cc.CurrentEngine;
            IEngineClient engineClient = EngineClientsPool.GetInstance().FromPool(ccEngine);



            var ccquery = CC.Current.WebServices.Get("trainingquery")?.AsQuery();

            BotQueryPlugin queryPlugin = new BotQueryPlugin();
            queryPlugin.userid=userid;
            queryPlugin.domain = this.domainName;

            JQuery jquery = new JQuery();


            //You can add a JsonMethodPlugin here to customize the behavior
            //jquery.Plugin=...
            jquery.CreateOrInitPlugin();


            //{"app":"training-search",
            //            "method": "queryintent",
            //"debug":"true",
            //"query": {
            //                "name": "trainingquery",
            //"text": "tower block"}
            //        }
            Json request = Json.NewObject();
            request.Set("app", "training-search");
            request.Set("method", "search");
            Json query = Json.NewObject();
            query.Set("name", "trainingquery");
            query.Set("text", text);
            query.Set("pageSize", 5);
            //query.Set("isFirstPage", true);
            request.Set("query", query);



            jquery.JsonRequest=request;
            jquery.JsonResponse = Json.NewObject();
            queryPlugin.InitExecute(jquery, ccquery);

            queryPlugin.Execute();

            //queryPlugin.InitExecute(jquery, ccquery);

            //queryPlugin.Init();

            //jquery.

            //jquery.Execute();

            //queryPlugin.DoQuery();

            Json records=jquery.JsonResponse.Get("records");

            if (records==null)
                return null;

            return  records.EnumerateElements().Select(
                record => (
                record.ValueStr("id"),
                record.ValueStr("authors"),
                record.ValueStr("smallsummaryhtml")?.Length > 1 ? record.ValueStr("smallsummaryhtml") : record.ValueStr("relevantExtracts"),
                record.ValueStr("modified"),
                record.ValueStr("treepath"),
                record.ValueStr("url1"),
                record.ValueStr("title"),
                record.ValueStr("thumbnailUrl"),
                record.ValueStr("objectType"),
                record.ValueStr("fileext"),
                text

                ));
            

            //var payload = $"{{\"app\": \"{ sinequa_app_name }\", \"query\": {{ \"name\": \"{sinequa_ws_query_name }\", \"text\": { JsonConvert.ToString(text) } , \"pageSize\" : {pageSize} }} }}";

            //return response["records"].Select(
            //item => (item["id"]?.ToString(),
            //item["authors"]?.ToString(),
            //item["smallsummaryhtml"]?.ToString().Length > 1 ? item["smallsummaryhtml"]?.ToString() : item["relevantExtracts"]?.ToString(),
            //item["modified"]?.ToString(),
            //item["treepath"]?.ToString(),
            //item["url1"]?.ToString(),
            //item["title"]?.ToString(),
            //item["thumbnailUrl"]?.ToString(),
            //item["objectType"]?.ToString(),
            //item["fileext"]?.ToString(),
            //text
            //));

            //return await FindResultsFromPayload(userid, payload, text);
        }


        private async Task<IEnumerable<(string, string, string, string, string, string, string, string, string, string, string)>> FindResultsFromPayload(string userid, string payload, string queryHint)
        {

            JObject jSonObject = await DoRequest(userid, payload, "query");
            return jSonObject["records"].Select(
                item => (item["id"]?.ToString(),
                item["authors"]?.ToString(),
                item["smallsummaryhtml"]?.ToString().Length > 1 ? item["smallsummaryhtml"]?.ToString() : item["relevantExtracts"]?.ToString(),
                item["modified"]?.ToString(),
                item["treepath"]?.ToString(),
                item["url1"]?.ToString(),
                item["title"]?.ToString(),
                item["thumbnailUrl"]?.ToString(),
                item["objectType"]?.ToString(),
                item["fileext"]?.ToString(),
                queryHint
                ));
        }

        private async Task<JObject> FindResultsFromPayloadEx(string userid, string payload, string queryHint)
        {
            return await DoRequest(userid, payload, "query");
        }



        private async Task<JObject> DoRequest(string userid, string payload, string verb)
        {
            var handler = new HttpClientHandler();
            //=============================================================================
            //Do not uncomment in a production environemnt, workaround for self signed certificate =========
            //handler.ClientCertificateOptions = ClientCertificateOption.Manual;
            //handler.ServerCertificateCustomValidationCallback =
            //    (httpRequestMessage, cert, cetChain, policyErrors) =>
            //    {
            //        return true;
            //    };
            //EOWorkaround ================================================================
            //=============================================================================
            var client = new HttpClient(handler);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", jwtBearerToken);
            //client.DefaultRequestHeaders.Add("X-On-Behalf-Of", $"{domainName}|{userid}");


            //building Sinequa endpoint URL (here SBA v2) 
            UriBuilder ub = new UriBuilder();
            ub.Scheme = "https";
            ub.Host = host;
            ub.Port = port;
            ub.Path = $"/api/v1/{verb}";

            _logger.LogDebug("Base URL = " + ub.Uri.AbsoluteUri + " Payload= " + payload);

            _logger.LogDebug($"jwtBearerToken = {jwtBearerToken.Substring(0, 5)}[...] X-On-Behalf-Of={domainName}|{ userid}");

            var endPoint = ub.Uri;
            HttpContent content = new StringContent(payload, Encoding.UTF8, "application/json");
            content.Headers.Add("X-On-Behalf-Of", $"{domainName}|{userid}");
            HttpResponseMessage res = null;
            JObject jSonObject = null;
            string strResponse = null;
            try
            {
                res = await client.PostAsync(endPoint, content);
                strResponse = await res?.Content?.ReadAsStringAsync();
                jSonObject = JObject.Parse(strResponse);
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Exception while posting request");
                if (!string.IsNullOrEmpty(strResponse))
                {
                    _logger.LogDebug($"Response content: [{strResponse}]");
                }
            }

            return jSonObject;
        }

        private async Task<string> GetBase64Thumbnail(string userid, string path)
        {
            var handler = new HttpClientHandler();
            //=============================================================================
            //Do not uncomment in a production environemnt, workaround for self signed certificate =========
            //handler.ClientCertificateOptions = ClientCertificateOption.Manual;
            //handler.ServerCertificateCustomValidationCallback =
            //    (httpRequestMessage, cert, cetChain, policyErrors) =>
            //    {
            //        return true;
            //    };
            //EOWorkaround ================================================================
            //=============================================================================
            var client = new HttpClient(handler);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", jwtBearerToken);
            client.DefaultRequestHeaders.Add("X-On-Behalf-Of", $"{domainName}|{userid}");

            //building Sinequa endpoint URL (here SBA v2) 
            UriBuilder ub = new UriBuilder();
            ub.Scheme = "https";
            ub.Host = host;
            ub.Port = port;
            ub.Path = path;

            var endPoint = ub.Uri;


            HttpResponseMessage res = null;
            byte[] bytes = null;
            var imageData = string.Empty;
            try
            {
                res = await client.GetAsync(endPoint);
                if (res == null || !res.IsSuccessStatusCode)
                {
                    return null;
                }
                bytes = await res?.Content?.ReadAsByteArrayAsync();
                imageData = Convert.ToBase64String(bytes);
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Exception while posting request");
                if (bytes == null || bytes.Length == 0)
                {
                    _logger.LogDebug($"Empty response]");
                }
            }
            return imageData;
        }
        private static string StripHtml(string source, bool highlightToMarkdown = true)
        {
            string cleanerText = WebUtility.HtmlDecode(source); //replace html entities 
            //regex is meant to replace <b> tags by their markdowm equivalent ** 
            //also addresses msft implementation issue with markdown (whitespaces prevent bold fonts to be rendered)
            return Regex.Replace(cleanerText, @"(<b>)([^<\s]+)(\s*</b>)", highlightToMarkdown ? "**$2** " : " ", RegexOptions.IgnoreCase | RegexOptions.ECMAScript);
        }

        /**
         * Used for formatting the queryText of some of the metadata field values
         */
        private static string StringCleanup(string inputString)
        {
            if (!string.IsNullOrWhiteSpace(inputString))
            {
                var outputString = inputString.Trim(new char[] { '[', ']', ' ', '\r', '\n', '\"' });
                return outputString.Replace("\"", "").Replace("\r", "").Replace("\n", "");
            }
            return string.Empty;
        }
        private static Attachment ImageData2InlineAttachment(string imageData)
        {
            return new Attachment
            {
                Name = $"Resources\\Thumbnail",
                ContentType = "image/jpeg",
                ContentUrl = $"data:image/jpeg;base64,{imageData}",
            };
        }

        private static Attachment Resource2InlineAttachment(string imageName, string contentType = null)
        {

            if (string.IsNullOrWhiteSpace(imageName))
            {
                throw new ArgumentNullException("imageName");
            }
            if (string.IsNullOrEmpty(contentType))
            {
                switch (Path.GetExtension(imageName).ToLower())
                {
                    case ".jpg":
                    case ".jpeg":
                    case ".jpfif":
                        contentType = "image/jpeg";
                        break;
                    case ".png":
                        contentType = "image/png";
                        break;
                    case ".gif":
                        contentType = "image/gif";
                        break;
                    case ".apng":
                        contentType = "image/apng";
                        break;
                    case ".svg":
                        contentType = "image/svg+xml";
                        break;
                    default:
                        contentType = "application/octet-stream";
                        break;
                }
            }
            var imagePath = Path.Combine(Environment.CurrentDirectory, @"Resources", imageName);
            var fileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("TeamsMessagingExtensionsSearch.Resources." + imageName);

            MemoryStream ms=new MemoryStream();
            fileStream.CopyTo(ms);
            var imageData = Convert.ToBase64String(ms.ToArray());

            //var imageData = Convert.ToBase64String(File.ReadAllBytes(imagePath));
            return new Attachment
            {
                Name = $"Resources\\{imageName}",
                ContentType = contentType,
                ContentUrl = $"data:{contentType};base64,{imageData}",
            };
        }

        private static string SourceFromTreePath(string treepath)
        {
            if (string.IsNullOrWhiteSpace(treepath))
            {
                return string.Empty;
            }
            string str = StripHtml(StringCleanup(treepath));
            string[] treepathArr = str.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (treepathArr != null && treepath.Length >= 1)
            {
                return StringCleanup(treepathArr[0]);
            }
            return string.Empty;
        }


        private static string ParseContent(string content, string ct)
        {
            //Possible values include: 'html', 'text'.
            string retValue = string.Empty;
            //if ("html".Equals(ct) && !string.IsNullOrWhiteSpace(content))
            //{
            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(content);
            retValue = document.DocumentNode.InnerText;
            //}
            //else if ("text".Equals(ct) && !string.IsNullOrWhiteSpace(content))
            //{
            //    retValue = content;
            //}
            return retValue;
        }


        //private MessagingExtensionActionResponse ShareMessageCommand(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
        //{
        //    // The user has chosen to share a message by choosing the 'Share Message' context menu command.
        //    var heroCard = new HeroCard
        //    {
        //        Title = $"{action.MessagePayload.From?.User?.DisplayName} orignally sent this message:",
        //        Text = action.MessagePayload.Body.Content,
        //    };

        //    if (action.MessagePayload.Attachments != null && action.MessagePayload.Attachments.Count > 0)
        //    {
        //        // This sample does not add the MessagePayload Attachments.  This is left as an
        //        // exercise for the user.
        //        heroCard.Subtitle = $"({action.MessagePayload.Attachments.Count} Attachments not included)";
        //    }

        //    // This Messaging Extension example allows the user to check a box to include an image with the
        //    // shared message.  This demonstrates sending custom parameters along with the message payload.
        //    var includeImage = ((JObject)action.Data)["includeImage"]?.ToString();
        //    if (string.Equals(includeImage, bool.TrueString, StringComparison.OrdinalIgnoreCase))
        //    {
        //        heroCard.Images = new List<CardImage>
        //        {
        //            new CardImage { Url = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU" },
        //        };
        //    }

        //    return new MessagingExtensionActionResponse
        //    {
        //        ComposeExtension = new MessagingExtensionResult
        //        {
        //            Type = "result",
        //            AttachmentLayout = "list",
        //            Attachments = new List<MessagingExtensionAttachment>()
        //            {
        //                new MessagingExtensionAttachment
        //                {
        //                    Content = heroCard,
        //                    ContentType = HeroCard.ContentType,
        //                    Preview = heroCard.ToAttachment(),
        //                },
        //            },
        //        },
        //    };
        //}

        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {



            var asJobject = JObject.FromObject(taskModuleRequest.Data);

            var previewTask = asJobject.ToObject<CardTaskFetchValue<PreviewTask>>()?.Data;


            var taskInfo = new TaskModuleTaskInfo();
            switch (previewTask.ActionId)
            {
                case TaskModuleIds.Preview:
                    //taskInfo.Url = taskInfo.FallbackUrl = $"{_baseUrl}/?#/preview?url={Convert.ToBase64String(Encoding.UTF8.GetBytes(previewTask.Url))}";
                    taskInfo.Url = taskInfo.FallbackUrl = previewTask.Url;
                    SetTaskInfo(taskInfo, TaskModuleUIConstants.Preview);
                    break;
                //More to be added ? 
                default:
                    break;
            }
            //var tmp = taskInfo.ToTaskModuleResponse();
            return await Task.FromResult(taskInfo.ToTaskModuleResponse());
        }

        //protected override Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        //{
        //    _logger.LogDebug("OnTeamsTaskModuleSubmitAsync called !");
        //    return null;
        //}

        //protected override Task<InvokeResponse> OnTeamsCardActionInvokeAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        //{
        //    _logger.LogDebug("OnTeamsCardActionInvokeAsync called !");
        //    return null;
        //}

        //   protected override Task OnTeamsMessagingExtensionCardButtonClickedAsync(ITurnContext<IInvokeActivity> turnContext, JObject cardData, CancellationToken cancellationToken)
        //{
        //    _logger.LogDebug("OnTeamsMessagingExtensionCardButtonClickedAsync called !");
        //    return null;
        //}

        private static void SetTaskInfo(TaskModuleTaskInfo taskInfo, UISettings uIConstants)
        {
            taskInfo.Height = uIConstants.Height;
            taskInfo.Width = uIConstants.Width;
            taskInfo.Title = uIConstants.Title.ToString();
        }


        protected override async  Task<InvokeResponse> OnInvokeActivityAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                if (turnContext.Activity.Name == SignInConstants.TokenExchangeOperationName && turnContext.Activity.ChannelId == Channels.Msteams)
                {
                    await OnTokenResponseEventAsync((ITurnContext <IEventActivity>) turnContext, cancellationToken);
                    return new InvokeResponse() { Status = 200 };
                }
                else
                {
                    return await base.OnInvokeActivityAsync(turnContext, cancellationToken);
                }
            }
            catch (InvokeResponseException e)
            {
                return e.CreateInvokeResponse();
            }
        }

        private async Task<TokenResponse> GetTokenResponse(ITurnContext<IInvokeActivity> turnContext, string state, CancellationToken cancellationToken)
        {
            var magicCode = string.Empty;

            if (!string.IsNullOrEmpty(state))
            {
                if (int.TryParse(state, out var parsed))
                {
                    magicCode = parsed.ToString();
                }
            }

            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
            var tokenResponse = await userTokenClient.GetUserTokenAsync(turnContext.Activity.From.Id, _connectionName, turnContext.Activity.ChannelId, magicCode, cancellationToken).ConfigureAwait(false);
            return tokenResponse;
        }

        private static async Task<HttpResponseMessage> PopUpSignInHandler(ITurnContext<IInvokeActivity> turnContext)
        {
            await turnContext.SendActivityAsync("Authentication Successful");
            return new HttpResponseMessage(HttpStatusCode.OK);
        }

    }
}
