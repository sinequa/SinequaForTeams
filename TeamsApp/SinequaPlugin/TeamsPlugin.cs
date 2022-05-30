using Microsoft.AspNetCore.Http;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sinequa.Microsoft.Teams;
using Sinequa.Microsoft.Teams.Bots;
using Sinequa.Plugins;
using System.Net.Http;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder;
using Microsoft.BotBuilderSamples;

namespace TeamsMessagingExtensionsSearch.SinequaPlugin
{


    public class TeamsPlugin : MessagingBotPlugin
    {
        //private readonly IBotFrameworkHttpAdapter Adapter;
        private TeamsMessagingExtensionsSearchBot<Dialog> Bot;
        private AdapterWithErrorHandler adapter;
        private IConfiguration configuration;
        private ILogger<TeamsMessagingExtensionsSearchBot<Dialog>> logger;

        public override void OnBegin()
        {
            System.Console.WriteLine(System.Environment.CurrentDirectory);

            var configurationBuilder = new ConfigurationBuilder()
           .AddJsonFile(@"C:\dev\git\ice\distrib\data\configuration\plugins\Teams\toto.json", optional: true, reloadOnChange: true);

            System.Console.WriteLine(configurationBuilder.Sources[0].ToString());

            //WARNING : appsettings.json is not properly loaded when built from sinequa admin or on server start.
            //TODO : FIXME
            configuration = configurationBuilder.Build();

            //In-file configuration for the time being...
            //Name of the Sinequa App
            configuration["Sinequa:AppName"] = "training-search";
            //Name of the query to use
            configuration["Sinequa:WSQueryName"] = "trainingquery";
            //Azure Active Directory used to match the teams User to Sinequa
            configuration["Sinequa:Domain"] = "AAD";

            //Teams Bot ID
            configuration["MicrosoftAppId"] = "e7340ec1-a658-4561-9161-f9e4998eeb11";// "6a15627b-8ddb-4f1f-8b99-a82b9417442e";
            //Teams Bot tenant ID
            configuration["MicrosoftAppTenantId"] = "2e572dab-8111-4c19-b7c3-439df4b8cd75";
            //Teams bot secret f any
            configuration["MicrosoftAppPassword"] = "CM18Q~O6wnPoIxRVahKVt8XVlgGc7O4P4Dwlacom";// "4838Q~HfWkrtk2xuOax-mhm-hye98m2wpC0kpaS_";//05a23ee8-6643-4809-8a20-c34664c0c7ae

            configuration["MicrosoftAppType"] = "MultiTenant";

            //For preview & direct link purposes.
            configuration["Sinequa:HostName"] = "7708-46-255-177-187.eu.ngrok.io";
            configuration["Sinequa:CustomPort"] = "";
            //configuration["Sinequa:BaseUrl"] = "https://localhost:444/app/training-search";
            configuration["Sinequa:BaseUrl"] = "https://7708-46-255-177-187.eu.ngrok.io";

            configuration["ConnectionName"] = "SSO Connection"; //Define on Azure App


            //Replace by an actual logger if needed
            ILogger botLogger = NullLogger<TeamsMessagingExtensionsSearchBot<Dialog>>.Instance;
            ILogger<MainDialog> dialogLogger = NullLogger<MainDialog>.Instance;

            MemoryStorage storage = new MemoryStorage();
            UserState userState = new UserState(storage);
            ConversationState conversationState = new ConversationState(storage);
            MainDialog dialog = new MainDialog(configuration, dialogLogger);


            logger = NullLogger<TeamsMessagingExtensionsSearchBot<Dialog>>.Instance;
            adapter = new AdapterWithErrorHandler(configuration);
            Bot = new TeamsMessagingExtensionsSearchBot<Dialog>(conversationState, userState, dialog, configuration, logger);
            isInitDone = true;

        }
        public override void OnPluginMethod(HttpContext context)
        {
            adapter.ProcessAsync(context.Request, context.Response, Bot);
        }
    }
}

