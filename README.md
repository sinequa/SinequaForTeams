# SinequaForTeams

---
**December 2022 Update**

This integration is not required anymore as the Teams integration is now part of the **11.9.0** Sinequa release.

This repository will be closed shortly.

---

# Sample Integration into Microsoft Teams

[Messaging Extensions](https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions) are a special kind of Microsoft Teams application that is support by the [Bot Framework](https://dev.botframework.com) v4.

This integration is an illustration of Search-based Messaging Extension (further reading on Messaging extensions : [Search-based](https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/search-commands/define-search-command) and [Action-based](https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/action-commands/define-action-command) )


## Prerequisites

- Office 365 account with Microsoft Teams  
- Azure subscription.
- Sinequa APIs endpoint (api/v1) over HTTPS that can be consumed from your messaging endpoint (from Azure App Service by default). 


## Sinequa Bot deployment using Terraform 

> Note these instructions are for deploying the bot on Azure infrastructure for a development environment.
> A change in the default SKUs (mostly Free tier or Dev tier) as well as the addition of extra resources for a more stringent security have to be considered before moving to a production environement. 

Here's a list of the various components created by the terraform script in its most comprehensive version:
- a Resource Group with the following resources 
  -  Bot Service (along with the Bot Channel Registration for Micorsoft Teams)
  -  App Service (and its App Service Plan)
  -  one Key Vault storing the following secrets: ApplicationID and Application Password created with the Application Registration, as well as the JWT used to authenticate the bot with the Sinequa platform. A default policy is assigned to the store only allowing the identity of the App Service instance to read secrets in the Key Vault)
-  The Application Registration for the current tenant.    


![Architecture](/Images/Teams-bot.png)


1. Preliminary steps:
   1. In **Sinequa admin console**, Open `Security / Access Tokens` and create a new `Bearer Token` for an administrator (required for impersonation) and provide access to the following JSon methods:
      - search.query
      - xdownload  
   2. Save the newly generated token, we'll refer to it later in this document as JWT
   3. You'll need a Sinequa security domain build on top of Azure Active Directory in the same tenant as your Teams/Office 365 (for authorization purpose) 
   4. **Identify** the following elements in the Search Based Application that you want to use as the main Search Service for your bot:
      - Fully Qualified Domain Name 
      - Application name
      - Web Service Name
      - The Port if different from 443
   5. In **Sinequa admin console**, `WebApps / <<Your Webapp>> `:
      - `Advanced` Tab: **Add** custom http headers on the webapp (for document preview in iframe): Content-Security-Policy : frame-ancestors teams.microsoft.com *.teams.microsoft.com *.skype.com *.ngrok.io *.azurewebsites.net
      - `Stateless Mode` Tab: **Add**  the value ``https://teams.microsoft.com`` to `Permitted origins for Cross-Origin Resource Sharing (CORS) requests`  


2. Clone the repository

    `bash
    git clone https://github.sinequa.com/Product/SinequaForTeams.git
    `

3. Use the [Terraform](https://learn.hashicorp.com/collections/terraform/azure-get-started) scripts provided to provision all resources required by the Teams Integration.
Two options here: 
   1. Create all the required Azure resources from the ground up 
   2. Leverage existing resources, for instance an existing App Service Plan or App service Environemnt,  a Key Vault or even leverage an existing resource group. These scripts will likely need to be adapted to your particular situation in 
   order to reuse and import one or more existing Azure resources.
  
Regardless of which option is chosen the output will contain:
- The value of the newly provisionned `Microsoft Application ID` (a.k.a. `Bot ID`), make sure to record this value, it will be needed in the following steps
- Information about the newly created resources such as: 
  - The URL of the newly created app service (default to https://<<app_svc_name>>.azurewebsites.net )
  - The resouce group name
  - ...

`
Apply complete! Resources: 15 added, 0 changed, 0 destroyed.

Outputs:

appservice_base_url = "https://myapplication.azurewebsites.net"
MicrosoftAppId = "123456ab-cd78-4f4f-ada2-c0b0l"
bot_channel_registration = {
[...]
}
resource_group =  {
[...]
}
`

>All the settings needed by the Sinequa teams Web Application are already configured, 

4. Update Microsoft Teams app package [Teams application manifest](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/apps-package)  
   1. **Edit** `manifest.json` located in the folder `TeamsAppManifest` Replace `<<bot_id>>` with the value of `MicrosoftAppId` obtained in the previous step (3 occurences)
   2. In `manifest.json`, **Append** your Sinequa Web application host name to the list of  `"validDomains"`
   3. Optional steps:  
      - Bring the Sinequa SBA experience into Teams by adding one or more [custom tabs](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/what-are-tabs) pointing to your existing SBAs .
      - Change the bot's short name (displayed to users in the Teams experience, default `SNQA` ) 
      - Change Sinequa Teams App Icons (Sinequa Logo by default)
   4. **Zip** up the contents of the `TeamsAppManifest` folder to create a `teamsappmanifest.zip` 
5. Upload your custom archive `teamsappmanifest.zip` to your teams app or subit it for approval. 
6. Deploy the bot implementation code to Azure ([Microsoft: Deploy your bot to Azure](https://aka.ms/deploy-your-bot))
   - Using Visual Studio 2019: Open the project and publish it to your App Service:
     - In Solution Explorer, right-click the project node and choose Publish (or use the Build > Publish menu item).
     - Select New and follow the wizard by selecting Azure , then the existing `resource group` (regardless if it was newly created or just reused in step 3) 
     followed by your existing instance of Azure App service and publish !
   - Using az cli: 
     - **Generate** a .deployment file in your bot project folder: 
     `d:\>[...]\TeamsApp>az bot prepare-deploy --lang Csharp --code-dir "." --proj-file-path "TeamsMessagingExtensionsSearch.csproj"`
     - **Manually create** a zip of the content <project-zip-path>  and **execute** the following command line from the project folder `d:\>[...]\TeamsApp>az webapp deployment source config-zip --resource-group "<resource-group-name>" --name "<name-of-web-app>" --src "<project-zip-path>"`
7.  Try it
    Once published if you don't have a broswer tab already open on it, navigate to `https://<<your app>>.azurewebsites.net` and you should see the following message:
   "Your Sinequa bot implementation has been provisionned! You can now test your bot in Teams." 
    

## Setting up SSO in tabs
  
Sinequa Search-Based Applications require users to be authenticated. Sinequa supports various standard Single-Sign-On protocols, like SAML or OAuth. However, these protocols involve a redirection to an Identity Provider (IdP) site, which the Teams clients does not accept.

The solution to this problem is to implement a different kind of SSO supported by Teams. This method is documented [here](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso). Concretely, it involves three main steps:
  1. Creating an Azure AD App registration, dedicated to the tab.
  2. Modifying the SBA running in the tab, so that it uses the Microsoft SDK to obtain a JSON Web Token (JWT) containing the user identity. Note that the SBA will still work as a standalone application with its own SSO; but in the context of Teams, it will use Teams' SSO.
  3. Adding a WebApp plugin to Sinequa to parse this token and extract the user identity.

Step 1 corresponds to the first and second part of the [Microsoft documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso?tabs=dotnet#1-create-your-azure-ad-application).

Step 2 consists in modyfing your SBA in the following way:
  - In your package.json file, add the following dependency: `"@microsoft/teams-js": "^1.10.0"` and run `npm install`.
  - Add the provided [teams-login.ts](https://github.com/sinequa/SinequaForTeams/blob/main/Tab-SSO/teams-login.ts) file to your SBA (next to `app.module.ts`)
  - Add the 2 following providers to your `app.module.ts` file:

```
{provide: AuthenticationService, useClass: TeamsAuthenticationService},
{provide: APP_INITIALIZER, useFactory: TeamsInitializer, deps: [AuthenticationService], multi: true},  
```
(where `APP_INITIALIZER` is imported from `@angular/common`, `AuthenticationService` from `@Sinequa/core/login` and the other ones from `./teams-login`)
  
  - Finally, recompile the app.

Step 3 consists in adding a [WebApp plugin](https://github.com/sinequa/SinequaForTeams/blob/main/Tab-SSO/TeamsWebappPlugin.cs) to the Sinequa WebApp hosting the SBA, so that it can validate the user identity. The attached plugin must be modified to match the specifics of your project. Particularly:
  - `sinequaDomain` is the security domain within which the user id provided by the SSO will be searched.
  - `sinequaAudience` corresponds to the app registration configured in Step 1 (in the form of `api://<Domain of the sinequa server>/<App registration id>`)
  - `sinequaIssuer` correponds to a tenant id for your company. If you have a doubt about this parameter, you can find it by decoding the JWT sent by Microsoft (using [jwt.io](https://jwt.io/)). The WebApp plugin just needs to check that the expected issuer matches the one in the JWT. (The plugin already logs the token and its issuer in the webapp logs).
  - `payloadIdField` corresponds to a field of the JWT containing the user identifier. For example, the default is `"upn"`, which is generally an email address. Whichever field of the JWT you choose, it must be consistent with the user ids in your Sinequa security domain.

## Further reading

- [How Microsoft Teams bots work](https://docs.microsoft.com/en-us/azure/bot-service/bot-builder-basics-teams?view=azure-bot-service-4.0&tabs=javascript)
- TODO Add pointer to sinequa howto , troubleshooting etc... 


## Disclaimer
All the code samples in here are for illustrative purposes only, they have not been fully tested. They are supplied “as is” without warranty of any kind, express or implied. SINEQUA assumes no responsibility or liability for the use of any code sample, conveys no license or title under any patent, copyright, or mask work right for the code sample.


