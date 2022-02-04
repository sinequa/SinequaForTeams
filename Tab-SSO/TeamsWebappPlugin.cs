///////////////////////////////////////////////////////////
// Plugin TeamsWebappPlugin : file TeamsWebappPlugin.cs
//

using System;
using Sinequa.Common;
using Sinequa.Plugins;
using Sinequa.Connectors;
//using Sinequa.Ml;

using System.IdentityModel.Tokens.Jwt;
using Microsoft.IdentityModel.Tokens;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Sinequa.Plugin
{
    public class TeamsWebappPlugin : WebAppPlugin
    {

        private static Json keys = null;
        private static string lastSecurityVersion;
        private static DateTime lastClearCache = DateTime.Now;

        private static readonly int refreshMinutes = 3*60; // 3 hour validity for the downloaded keys
        private static readonly string teamsJWTHeader = "teams-token"; //internal name for o365 token 
        private static readonly string microsoftKeysUrl = "https://login.microsoftonline.com/common/discovery/keys";
        private static readonly string sinequaDomain = "Sinequa";
        private static readonly string sinequaAudience = "api://testing.sinequa.com/2f80dff2-76f9-4b15-871c-befb7ac15846"; //actual app registration for SSO
        private static readonly string sinequaIssuer = "https://sts.windows.net/465ec3fd-500e-4e38-a426-5ca3086440bd/"; //tenant id
        private static readonly string payloadIdField = "upn";

        public override LoginInfo GetLoginInfo(IDocContext ctxt)
        {
            Sys.Log("TeamsWebappPlugin GetLoginInfo start");
            string token = ctxt.Hm.RequestHeader(teamsJWTHeader);
            if(!Str.IsEmpty(token)) {
                Sys.Log("TeamsWebappPlugin GetLoginInfo JWT: " + token);
                var jwt = Validate(token, ctxt.Cc.SecurityVersion);
                if(jwt != null) {
                    var id = (string)jwt.Payload[payloadIdField];
                    Sys.Log("User id: ", id);
                    var info = new LoginInfo();
                    info.DomainName = sinequaDomain;
                    info.UserName = id;
                    return info;
                }
            }
            Sys.Log("TeamsWebappPlugin GetLoginInfo end");
            return base.GetLoginInfo(ctxt);
        }

        public static JwtSecurityToken Validate(string token, string securityVersion)
        {
            var handler = new JwtSecurityTokenHandler();
            var decodedToken = handler.ReadJwtToken(token) as JwtSecurityToken;
            
            // Get the key ID from the header part  https://datatracker.ietf.org/doc/html/rfc7515
            // Example : nOo3ZDrODXEK1jKWhXslHR_KXEg
            // In order to seek out the publicKey
            string kid = (string)decodedToken.Header["kid"];
            
            Sys.Log("Issuer from jwt: "+ (string)decodedToken.Payload["iss"]);
            Sys.Log("Issuer expected: "+ sinequaIssuer);

            string keysAsString = null;
            // Discover the Azure Active Directory Key signatures
            try
            {
                // Check cached keys validity
                if (lastSecurityVersion != securityVersion 
                   || DateTime.Now.Subtract(lastClearCache).Minutes >= refreshMinutes)
                {
                    lastSecurityVersion = securityVersion;
                    lastClearCache = DateTime.Now;
                    keys = null;
                    Sys.Log("Azure keys cache cleared");
                }
                
                if(keys == null)
                {
                    Sys.Log($"GET {microsoftKeysUrl} ...");
                    var azureKeys = new UrlAccess().GetJson(microsoftKeysUrl);
                    if(azureKeys == null)
                    {
                        throw new Exception($"Failed to get Azure keys from {microsoftKeysUrl}");
                    }
                    Sys.Log($"Success: {Json.Serialize(azureKeys)}");
                    // Get the appropriate publicKey by the "kid"
                    keys = azureKeys.GetAsArray("keys");
                    if(keys == null)
                    {
                        throw new Exception($"Expected \"keys\" property from {microsoftKeysUrl}");
                    }
                }
                
                // Search from the key with the right 'kid'
                Json signatureKeyIdentifier = null;
                for(int i=0; i<keys.EltCount(); i++) {
                    if(keys.Elt(i).ValueStr("kid") == kid) {
                        signatureKeyIdentifier = keys.Elt(i); //.keys.FirstOrDefault(key => key.kid.Equals(kid));
                        break;
                    }
                }
                if (signatureKeyIdentifier != null)
                {
                    // Get the public Key from the http's response
                    string signatureKey = signatureKeyIdentifier.GetAsArray("x5c").EltStr(0);

                    // Uncomment the line below if you want more information in case of error
                    // IdentityModelEventSource.ShowPII = true;

                    // Create a X509 Certificate in order to create an RsaSecurityKey needed
                    // for the token's validation
                    var certificate = new X509Certificate2(Convert.FromBase64String(signatureKey));
                    RSA rSA = certificate.PublicKey.Key as RSA;
                    TokenValidationParameters validationParameters = new TokenValidationParameters
                    {
                        // This particular audience is the Azure Active Directory application audience
                        // for the SSO with Teams
                        // So only JWT with this audience will be validate and no other one
                        ValidateAudience = true,
                        ValidAudience = sinequaAudience,
                        // In this case, only these two issuers are allowed to access the application
                        // Here you have to populate dynamicly a list of issuers, depending the numbers
                        // of clients you have authorized to access your app i.e (multi-tenant app)
                        ValidateIssuer = true,
                        ValidIssuer = sinequaIssuer,
                        // Don't forget to set ValidateLifeTime to true in production
                        ValidateLifetime = false,
                        // Without this key, we aren't be able to validate the JWT
                        IssuerSigningKey = new RsaSecurityKey(rSA)
                    };
                    SecurityToken jwt;
                    var result = handler.ValidateToken(token, validationParameters, out jwt);
                    return jwt as JwtSecurityToken;
                }
                else
                {
                    Sys.Log($"Error: No key with kid={kid} from {microsoftKeysUrl}");
                }
                
                //return decodedToken;

            }
            catch (Exception ex)
            {
                Sys.Log("Error in Validation of Teams JWT: "+ex.Message);
            }
            return null;

        }



    }
}