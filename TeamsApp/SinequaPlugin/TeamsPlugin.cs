using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Sinequa.Microsoft.Teams;
using Sinequa.Plugins;
using Microsoft.Bot.Builder;
using Sinequa.Common;
using Microsoft.Bot.Builder.Teams;
using Sinequa.Web.TeamsBot;
using System;
using Sinequa.Search.JsonMethods;
using Sinequa.Configuration;
using System.Net;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Threading.Tasks;
using Microsoft.Bot.Schema.Teams;
using Newtonsoft.Json.Linq;
using Microsoft.Bot.Schema;
using System.Threading;
using System.Linq;
using System.Collections.Generic;
using AdaptiveCards;
using Microsoft.AspNetCore.WebUtilities;

namespace TeamsMessagingExtensionsSearch.SinequaPlugin
{

    public class TeamsPlugin : MessagingBotPlugin
    {

        public override TeamsActivityHandler GetTeamsActivityHandler()
        {
            return new MyTeamActivityHandler();
        }

        //This sample class does exactly the same as the standard behavior
        //All methods do exactly the same as the ones from the base class.
        //You do not need to override methods that you do not modify
        public class MyTeamActivityHandler : SinequaTeamsActivityHandler
        {
            #region utility functions
            //Searches a document by its ID.
            protected override Json FindDocumentById(string userid, string text, TeamsAdapter adapter, string id)
            {
                var ccquery = CC.Current.WebServices.Get(adapter.SinequaQueryName)?.AsQuery();

                if (!SetSessionUser(userid))
                    throw new Exception($"Could not set Session user with userid = {userid}");
                Json query = Json.NewObject();
                query.Set("name", adapter.SinequaQueryName);
                query.Set("text", text);
                Session.CurrentApp = CC.Current.AllApps.GetApp(adapter.SinequaAppName);
                var previewRecord = QueryPlugin.GetPreviewRecord(id, query, ccquery, Session);
                return previewRecord;
            }

            //Searches documents matching the "text" as input
            protected override Json FindResults(string userid, string text, TeamsAdapter adapter, out int resultCount, int pageSize = 5)
            {
                if (!SetSessionUser(userid))
                    throw new Exception($"Could not set Session user with userid = {userid}");

                Json query = Json.NewObject();
                query.Set("name", adapter.SinequaQueryName);
                query.Set("text", text);
                query.Set("pageSize", pageSize);
                var jquery = JQuery.NewQuery(Session, query);
                jquery.JsonRequest.Set("app", adapter.SinequaAppName);
                jquery.JsonRequest.Set("method", "search");
                jquery.Execute();

                resultCount = jquery.JsonResponse.ValueInt("totalRowCount");
                return jquery.JsonResponse.Get("records");
            }
            private static string StripHtml(string source, bool highlightToMarkdown = true)
            {
                if (string.IsNullOrEmpty(source))
                    return null;
                string cleanerText = WebUtility.HtmlDecode(source); //replace html entities 
                                                                    //regex is meant to replace <b> tags by their markdowm equivalent ** 
                                                                    //also addresses msft implementation issue with markdown (whitespaces prevent bold fonts to be rendered)
                return Regex.Replace(cleanerText, @"(<b>)([^<\s]+)(\s*</b>)", highlightToMarkdown ? "**$2** " : " ", RegexOptions.IgnoreCase | RegexOptions.ECMAScript);
            }
            private static string StringCleanup(string inputString)
            {
                if (!string.IsNullOrWhiteSpace(inputString))
                {
                    var outputString = inputString.Trim(new char[] { '[', ']', ' ', '\r', '\n', '\"' });
                    return outputString.Replace("\"", "").Replace("\r", "").Replace("\n", "");
                }
                return string.Empty;
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

            protected override Uri GetDirectLinkURL(string documentId, string queryText, TeamsAdapter adapter)
            {
                var urlEncodedID = System.Net.WebUtility.UrlEncode(documentId);
                var urlEncodedJSonPayload = System.Net.WebUtility.UrlEncode($"{{\"name\":{ JsonConvert.SerializeObject(adapter.SinequaQueryName)},\"text\":{JsonConvert.SerializeObject(queryText)}}}");
                UriBuilder ub = new UriBuilder();
                ub.Scheme = "https";
                ub.Host = adapter.Host;
                ub.Port = adapter.Port;
                ub.Path = $"/app/{adapter.SinequaAppName}/";
                //workaround for the sharp sign (otherwise urlencoded by uribuilder), tried the fragment part  w/o success...
                ub.Query = $"#/preview?id={urlEncodedID}&query={urlEncodedJSonPayload}";
                return ub.Uri;
            }

            private JObject RecordToJObject(Json record, string text)
            {
                return JObject.FromObject((
                    record.ValueStr("id"),
                    record.ValueStr("authors"),
                    record.ValueStr("smallsummaryhtml")?.Length > 1 ? record.ValueStr("smallsummaryhtml") : record.ValueStr("relevantExtracts"),
                    record.ValueDat("modified"),
                    record.ValueStr("treepath"),
                    record.ValueStr("url1"),
                    record.ValueStr("title"),
                    record.ValueStr("thumbnailUrl"),
                    record.ValueStr("objectType"),
                    record.ValueStr("fileext"),
                    text));
            }

            #endregion

            #region activity handling

            //Method called when someone is added to a group chat where the bot is active
            protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
            {
                var welcomeText = "Hello and welcome!";
                foreach (var member in membersAdded)
                {
                    if (member.Id != turnContext.Activity.Recipient.Id)
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText, welcomeText), cancellationToken).ConfigureAwait(false);
                    }
                }
            }

            //Method called when a message is sent in the bot chat
            protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
            {

                var adapter = (TeamsAdapter)turnContext.Adapter;
                if (turnContext.Activity.Value != null)
                {
                    var obj = (JObject)turnContext.Activity.Value;
                    var answer = obj["Answer"]?.ToString();
                }
                else //regular text message
                {
                    var queryText = turnContext?.Activity?.Text;

                    if (String.IsNullOrWhiteSpace(queryText))
                    {
                        if (turnContext.Activity?.Attachments?.Count >= 1)
                        {
                            await turnContext.SendActivityAsync("Sorry, I can only process queries from simple text. I can't handle attachments such as cards :(").ConfigureAwait(false);
                        }
                        else
                        {
                            await turnContext.SendActivityAsync("Please enter a valid search query (Empty search not availaible from Teams Sinequa App)").ConfigureAwait(false);
                        }
                        return;
                    }

                    var user = turnContext?.Activity.From;

                    Lg.Plugin.Trace("Teams UserID  = " + user.Id + "Teams Username :user.Name = " + user.Name + " Azure AD ID= " + user.AadObjectId);
                    try
                    {
                        var records = FindResults(adapter.DomainName + "|" + user.AadObjectId, queryText, (TeamsAdapter)turnContext.Adapter, out int resultCount);

                        // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
                        // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.
                        if (records == null)
                        {
                            await turnContext.SendActivityAsync("Search error (is engine started ?)").ConfigureAwait(false);
                            return;
                        }



                        var attachments = records.EnumerateElements().Select(record =>
                        {
                            var template = adapter.Templates["recordDisplayCard"];

                            var docDirectLinkUrl = GetDirectLinkURL(record.ValueStr("id"), queryText, (TeamsAdapter)turnContext.Adapter)?.AbsoluteUri;
                            DateTime modifiedDT = record.ValueDat("modified", DateTime.Now);

                            if (!adapter.Attachments.TryGetValue(record.ValueStr("fileext"), out var image))
                                adapter.Attachments.TryGetValue("any", out image);

                            var myData = new
                            {
                                DocTitle = record.ValueStr("title"),
                                Summary = StripHtml(record.ValueStr("smallsummaryhtml")?.Length > 1 ? record.ValueStr("smallsummaryhtml") : record.ValueStr("relevantExtracts")), // ConvertHTMLToMarkdown(relevantExtracts),  //<b>...</b>  //**...**  //encoding issue on author
                                ThumbnailUrl = image?.ContentUrl,
                                PreviewUrl = docDirectLinkUrl,
                                DirectlinkUrl = docDirectLinkUrl,
                                SourceTreepath = SourceFromTreePath(record.ValueStr("treepath")),
                                AuthorName = StringCleanup(record.ValueStr("authors")),  //encoding issue 
                                FileType = record.ValueStr("objectType"),
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

                        await turnContext.SendActivityAsync(MessageFactory.Carousel(attachments, attachments?.Count > 0 ? $"I found {resultCount} result(s). Top 5 results :\r\n" : "Your search did not match any documents"), cancellationToken).ConfigureAwait(false);
                    }
                    catch (Exception)
                    {
                        await turnContext.SendActivityAsync("Search error. This might be because the search engine couldn't map your Teams user ID to a valid Sinequa user").ConfigureAwait(false);
                        throw;
                    }
                }

            }

            //Method called on fetch activities. For example the "Preview button" defined in the sample cards
            protected override Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
            {
                JObject taskModuleRequestData = (JObject)taskModuleRequest.Data;

                var url = taskModuleRequestData["data"]["url"].ToString();
                var taskInfo = new TaskModuleTaskInfo();
                switch (taskModuleRequestData["data"]["actionid"].ToString())
                {
                    case "preview":
                        taskInfo.Url = taskInfo.FallbackUrl = url;
                        taskInfo.Width = 1000;
                        taskInfo.Height = 700;
                        taskInfo.Title = "Document Preview";
                        break;
                    //More to be added ? 
                    default:
                        break;
                }
                return Task.FromResult(new TaskModuleResponse
                {
                    Task = new TaskModuleContinueResponse()
                    {
                        Value = taskInfo,
                    },
                });
            }
            //Method called by the messaging extension (Search bar and chossing the app from the message field in a chat)
            protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
            {
                var adapter = (TeamsAdapter)turnContext.Adapter;
                var queryText = query?.Parameters?[0]?.Value as string ?? string.Empty;
                var user = turnContext?.Activity.From;

                Lg.Plugin.Debug("Teams UserID  = " + user.Id + "Teams Username :user.Name = " + user.Name + " Azure AD ID= " + user.AadObjectId);

                var records = FindResults(adapter.DomainName + "|" + user.AadObjectId, queryText, (TeamsAdapter)turnContext.Adapter, out var resultCount);

                //Response from FindResults- packages=> id, authors, smallsummaryhtml, modified, treepath, url1, title, thumbnail, objectType, fileext,queryText

                // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
                // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.
                var attachments = records.EnumerateElements().Select(record =>
                {
                    DateTime modifiedDT = record.ValueDat("modified", DateTime.Now);
                    string strModified = modifiedDT.ToString("ddd, MMM d yyyy");

                    var previewCard = new ThumbnailCard
                    {
                        Title = record.ValueStr("title"),
                        Subtitle = $"{StringCleanup(record.ValueStr("authors"))} - {strModified}",
                        Text = StripHtml(record.ValueStr("smallsummaryhtml")?.Length > 1 ? record.ValueStr("smallsummaryhtml") : record.ValueStr("relevantExtracts"), false),
                        Tap = new CardAction { Type = "invoke", Value = RecordToJObject(record, queryText) }
                    };

                    if (!adapter.Attachments.TryGetValue(record.ValueStr("fileext"), out var image))
                        adapter.Attachments.TryGetValue("any", out image);
                    if (image != null)
                        previewCard.Images = new List<CardImage>() { new CardImage(image.ContentUrl, record.ValueStr("url1"), "OpenUrl") };


                    var attachment = new MessagingExtensionAttachment
                    {
                        ContentType = HeroCard.ContentType,
                        Content = new HeroCard { Title = record.ValueStr("title"), Text = record.ValueStr("smallsummaryhtml")?.Length > 1 ? record.ValueStr("smallsummaryhtml") : record.ValueStr("relevantExtracts") },
                        Preview = previewCard.ToAttachment()
                    };
                    return attachment;
                }).ToList();

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

            //Method called when selecting a result from the messaging extension
            protected override Task<MessagingExtensionResponse> OnTeamsMessagingExtensionSelectItemAsync(ITurnContext<IInvokeActivity> turnContext, JObject query, CancellationToken cancellationToken)
            {

                var adapter = (TeamsAdapter)turnContext.Adapter;
                var user = turnContext?.Activity.From;

                // The Preview card's Tap should have a Value property assigned, this will be returned to the bot in this event. 
                var (id, authors, relevantExtracts, modified, treepath, url1, title, thumbnailUrl, objectType, fileext, text) = query.ToObject<(string, string, string, string, string, string, string, string, string, string, string)>();
                DateTime modifiedDT = DateTime.Parse(modified);
                var docCacheURL = GetDirectLinkURL(id, text, adapter)?.AbsoluteUri;
                Lg.Plugin.Debug("Doc Cache URL = " + docCacheURL.ToString());

                if (!adapter.Attachments.TryGetValue(fileext, out var image))
                    adapter.Attachments.TryGetValue("any", out image);

                if (treepath != null && treepath.Length > 0)
                    treepath = StringCleanup(treepath);

                if (authors != null && authors.Length > 0)
                    authors = StringCleanup(authors);

                adapter.Templates.TryGetValue("recordDisplayCard", out var template);

                //Convert Summary content from HTML to Markdown as needed by Adaptive Cards
                var myData = new
                {
                    DocTitle = title,
                    Summary = StripHtml(relevantExtracts), // ConvertHTMLToMarkdown(relevantExtracts),  //<b>...</b>  //**...**  //encoding issue on author
                    ThumbnailUrl = image.ContentUrl,
                    PreviewUrl = GetDirectLinkURL(id, text, adapter)?.AbsoluteUri,
                    DirectlinkUrl = url1,
                    SourceTreepath = treepath,
                    AuthorName = StringCleanup(authors),  //encoding issue 
                    FileType = fileext,
                    Modified = modifiedDT.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ssK")
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

                Lg.Plugin.Trace("adaptiveCardattachment Content = " + adaptiveCardattachment.Content.ToString());

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

            //Method called when writing a URL in the message field that 
            protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
            {
                var adapter = (TeamsAdapter)turnContext.Adapter;
                try
                {
                    if (query == null)
                        throw new ArgumentNullException("query");
                    if (turnContext == null)
                        throw new ArgumentNullException("turnContext");

                    var user = turnContext.Activity?.From;
                    Uri uri = new Uri(query.Url);

                    if (uri != null && !string.IsNullOrWhiteSpace(uri.AbsolutePath))
                    {
                        if (!string.IsNullOrEmpty(uri.Fragment)
                            && uri.AbsolutePath.Equals($"/app/{adapter.SinequaAppName}/", StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (uri.Fragment.StartsWith("#/search?query=", StringComparison.InvariantCultureIgnoreCase))
                            {
                                string jsonPayload = uri.Fragment.Substring("#/search?query=".Length);
                                jsonPayload = WebUtility.UrlDecode(jsonPayload);

                                if (jsonPayload.Length > 0)
                                {
                                    JObject obj = JObject.Parse(jsonPayload);

                                    var queryTextHint = obj?.GetValue("text")?.Value<String>();

                                    Json rootObject = FindResults(adapter.DomainName + "|" + user.AadObjectId, queryTextHint, adapter, out int resCount);
                                    adapter.Attachments.TryGetValue("logo", out var img);

                                    MessagingExtensionAttachment attachment;
                                    if (rootObject.EltCount() == 0)
                                    {

                                        var heroCard2 = new ThumbnailCard
                                        {
                                            Title = $"Query \"{queryTextHint}\" : No document found ",
                                            Text = $"Search matched no result",
                                            Images = new List<CardImage> {
                                                new CardImage(img.ContentUrl, "", null),
                                    },
                                            Tap = new CardAction { Type = "invoke", Value = null }
                                        };

                                        attachment = new MessagingExtensionAttachment(ThumbnailCard.ContentType, null, content: heroCard2, preview: heroCard2.ToAttachment());
                                    }
                                    else
                                    {
                                        var template = adapter.Templates["unfurlSearch"];
                                        if (!adapter.Attachments.TryGetValue("logo32", out var image))
                                            adapter.Attachments.TryGetValue("any", out image);

                                        var myData = new
                                        {
                                            Title = $"Sinequa search : \"{queryTextHint}\"",
                                            LogoUrl = image?.ContentUrl,
                                            PreviewUrl = uri,
                                            DirectlinkUrl = uri,
                                            SearchDescription = $"{resCount} result(s) found"
                                        };
                                        var cardJson = template.Expand(myData);
                                        var adaptiveCardAttachment = new Attachment()
                                        {
                                            ContentType = "application/vnd.microsoft.card.adaptive",
                                            Content = JsonConvert.DeserializeObject(cardJson),
                                        };

                                        attachment = new MessagingExtensionAttachment(AdaptiveCard.ContentType, null, content: JsonConvert.DeserializeObject(cardJson), preview: adaptiveCardAttachment);
                                    }
                                    var msgngExtRslt = new MessagingExtensionResult(AttachmentLayoutTypes.List, "result", attachments: new[] { attachment }, text: $"Results for query : {queryTextHint}");
                                    return new MessagingExtensionResponse(msgngExtRslt);

                                }
                            }
                            else if (uri.Fragment.StartsWith("#/preview?", StringComparison.InvariantCultureIgnoreCase))
                            {
                                string decodedFragment = System.Web.HttpUtility.HtmlDecode(uri.Fragment);
                                var queryParameters = QueryHelpers.ParseQuery(decodedFragment.Substring("#/preview?".Length));
                                var docid = queryParameters?["id"];
                                MessagingExtensionAttachment attachment;
                                if (!string.IsNullOrWhiteSpace(docid))
                                {
                                    var queryDefJson = queryParameters?["query"];
                                    JObject obj = JObject.Parse(queryDefJson);
                                    var queryTextHint = obj?.GetValue("text")?.Value<String>();
                                    Json record = FindDocumentById(adapter.DomainName + "|" + user.AadObjectId, queryTextHint, adapter, docid);
                                    if (record != null)
                                    {
                                        DateTime modifiedDT = DateTime.Now;
                                        if (!DateTime.TryParse(record.ValueStr("modified"), out modifiedDT))
                                        {
                                            modifiedDT = DateTime.Now;
                                        }
                                        var template = adapter.Templates["recordDisplayCard"];

                                        var docDirectLinkUrl = GetDirectLinkURL(record.ValueStr("id"), queryTextHint, (TeamsAdapter)turnContext.Adapter)?.AbsoluteUri;

                                        if (!adapter.Attachments.TryGetValue(record.ValueStr("fileext"), out var image))
                                            adapter.Attachments.TryGetValue("any", out image);

                                        var myData = new
                                        {
                                            DocTitle = record.ValueStr("title"),
                                            Summary = StripHtml(record.ValueStr("smallsummaryhtml")?.Length > 1 ? record.ValueStr("smallsummaryhtml") : record.ValueStr("relevantExtracts")), // ConvertHTMLToMarkdown(relevantExtracts),  //<b>...</b>  //**...**  //encoding issue on author
                                            ThumbnailUrl = image?.ContentUrl,
                                            PreviewUrl = docDirectLinkUrl,
                                            DirectlinkUrl = docDirectLinkUrl,
                                            SourceTreepath = SourceFromTreePath(record.ValueStr("treepath")),
                                            AuthorName = StringCleanup(record.ValueStr("authors")),  //encoding issue 
                                            FileType = record.ValueStr("objectType"),
                                            Modified = modifiedDT.ToUniversalTime().ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ssK")
                                        };
                                        var cardJson = template.Expand(myData);
                                        var adaptiveCardAttachment = new Attachment()
                                        {
                                            ContentType = "application/vnd.microsoft.card.adaptive",
                                            Content = JsonConvert.DeserializeObject(cardJson),
                                        };
                                        attachment = new MessagingExtensionAttachment(AdaptiveCard.ContentType, null, content: JsonConvert.DeserializeObject(cardJson), preview: adaptiveCardAttachment);
                                    }
                                    else
                                    {
                                        adapter.Attachments.TryGetValue("logo", out var img);
                                        var thumbNailCard = new ThumbnailCard
                                        {
                                            Title = $"Oops, File not found",
                                            Text = $"This might be due to unsufficent user permissions",
                                            Images = new List<CardImage> { new CardImage(img.ContentUrl) },
                                            Tap = new CardAction { Type = "invoke", Value = null }
                                        };
                                        attachment = new MessagingExtensionAttachment(ThumbnailCard.ContentType, null, content: thumbNailCard, preview: thumbNailCard.ToAttachment());
                                    }

                                    var msgngExtRslt = new MessagingExtensionResult("list", "result", attachments: new[] { attachment }, text: "Document Preview");

                                    return new MessagingExtensionResponse(msgngExtRslt);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //Invalid sinequa URL will cause search errors
                    Lg.Plugin.Trace("Exception during link unfurling - query = " + query + " query.url = " + query.Url + "\n" + ex.Message);
                    var heroCardException = new ThumbnailCard
                    {
                        Title = "Oops...",
                        Text = $"An exception occured when trying to unfurl this link : {query.Url}]",
                        Tap = new CardAction { Type = "invoke", Value = query.Url }
                    };

                    var messagingExtensionAttEx = new MessagingExtensionAttachment(ThumbnailCard.ContentType, null, heroCardException);
                    var resultEx = new MessagingExtensionResult("list", "result", new[] { messagingExtensionAttEx });
                    return new MessagingExtensionResponse(resultEx);
                }

                Lg.Plugin.Trace("Unknown link unfurling - query = " + query + " query.url = " + query.Url);
                var heroCard = new ThumbnailCard
                {
                    Title = "Oops...",
                    Text = $"We don't know how to handle this link  {query.Url}]",
                    Tap = new CardAction { Type = "invoke", Value = query.Url }
                };

                var messagingExtensionAtt = new MessagingExtensionAttachment(ThumbnailCard.ContentType, null, heroCard);
                var result = new MessagingExtensionResult("list", "result", new[] { messagingExtensionAtt });
                return new MessagingExtensionResponse(result);
            }
            #endregion
        }
    }
}

