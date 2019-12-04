using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;

namespace SiteDesignTemplate.Provision
{
    class SiteDesign
    {
        public Guid siteDesignId;
    }
    class Metadata
    {
        public string type { get; set; }
    }
    class Metadata2
    {
        public string type { get; set; }
    }
    class SupportedSchemaVersions
    {
        public Metadata2 __metadata { get; set; }
        public List<string> results { get; set; }
    }
    class GetContextWebInformation
    {
        public Metadata __metadata { get; set; }
        public int FormDigestTimeoutSeconds { get; set; }
        public string FormDigestValue { get; set; }
        public string LibraryVersion { get; set; }
        public string SiteFullUrl { get; set; }
        public SupportedSchemaVersions SupportedSchemaVersions { get; set; }
        public string WebFullUrl { get; set; }
    }
    class D
    {
        public GetContextWebInformation GetContextWebInformation { get; set; }
    }
    class RootObject
    {
        public D d { get; set; }
    }
    public static class SiteDesignTemplate
    {
        public static bool ApplySiteDesign(ClientContext context, Guid siteDesignGUID)
        {

            bool isSiteDesignApplied = false;
            using (HttpClientHandler handler = new HttpClientHandler())
            {
                // Set permission setup accordingly for the call
                handler.Credentials = context.Credentials;
                handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));

                try
                {
                    using (HttpClient httpClient = new HttpClient(handler))
                    {
                        //POST
                        string requestUrl = string.Format("{0}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.AddSiteDesignTaskToCurrentWeb", context.Web.Url);

                        // Serialize request object to JSON
                        SiteDesign siteDesign = new SiteDesign
                        {
                            siteDesignId = siteDesignGUID
                        };

                        string jsonModernSite = JsonConvert.SerializeObject(siteDesign);
                        HttpContent body = new StringContent(jsonModernSite);

                        // Build Http request
                        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                        request.Headers.Add("accept", "application/json;odata=verbose");
                        MediaTypeHeaderValue.TryParse("application/json;odata=verbose;charset=utf-8", out MediaTypeHeaderValue sharePointJsonMediaType);
                        body.Headers.ContentType = sharePointJsonMediaType;

                        // Get Request Digest needed for post operation
                        string digestTask = GetRequestDigest(context);

                        // Deserialize the Request Digest data for getting formDigestValue
                        JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings
                        {
                            MissingMemberHandling = MissingMemberHandling.Ignore
                        };
                        RootObject contextInformation = JsonConvert.DeserializeObject<RootObject>(digestTask, jsonSerializerSettings);

                        // Add rest of the needed hearders
                        string formDigestValue = contextInformation.d.GetContextWebInformation.FormDigestValue;
                        body.Headers.Add("odata-version", "4.0");
                        body.Headers.Add("X-RequestDigest", formDigestValue);

                        // Perform actual post operation
                        HttpResponseMessage response = httpClient.PostAsync(requestUrl, body).Result;
                        isSiteDesignApplied = response.IsSuccessStatusCode;
                    }
                }
                catch (Exception ex)
                {
                    throw;
                }
                // Return response string to caller
                return isSiteDesignApplied;
            }
        }

        private static string GetRequestDigest(ClientContext context)
        {
            using (HttpClientHandler handler = new HttpClientHandler())
            {
                string responseString = string.Empty;

                handler.Credentials = context.Credentials;
                handler.CookieContainer.SetCookies(new Uri(context.Web.Url), (context.Credentials as SharePointOnlineCredentials).GetAuthenticationCookie(new Uri(context.Web.Url)));

                using (HttpClient httpClient = new HttpClient(handler))
                {
                    string requestUrl = string.Format("{0}/_api/contextinfo", context.Web.Url);
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=verbose");

                    HttpResponseMessage response = httpClient.SendAsync(request).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        responseString = response.Content.ReadAsStringAsync().Result;
                    }
                    else
                    {
                        throw new Exception(response.Content.ReadAsStringAsync().Result);
                    }
                }
                return responseString;
            }
        }

        public static void ApplyPnpTemplate(Web web)
        {
            try
            {
              
                string schemaDir = @"D:\GitHub\SPFx\SiteDesignTemplate\SiteDesignTemplate";
                //XMLTemplateProvider sitesProvider = new XMLFileSystemTemplateProvider(schemaDir, "");
                XMLTemplateProvider sitesProvider = new XMLFileSystemTemplateProvider(schemaDir, "");
                Console.WriteLine("Applying Template to Create the Project site Information List in Modern Sharepoint Site..." + schemaDir);
                ProvisioningTemplate template = sitesProvider.GetTemplate("PnPSiteTemplate.xml");
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                {
                    ProgressDelegate = (message, progress, total) =>
                    {
                        Console.WriteLine(string.Format("{0:00}/{1:00} - {2}", progress, total, message));
                    }
                };
                web.ApplyProvisioningTemplate(template, ptai);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: "+ ex.Message);
            }
        }
    }
}
