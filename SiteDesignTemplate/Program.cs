using Microsoft.SharePoint.Client;
using System;
using System.Net;

namespace SiteDesignTemplate
{
    class Program
    {
        static void Main(string[] args)
        {

            using (ClientContext ctx = new ClientContext("https://anibapatdev.sharepoint.com/sites/TTPrivateChannelSite-FirstPrivate"))
            {
                ctx.Credentials = new SharePointOnlineCredentials("campus-teams@anibapatdev.onmicrosoft.com", new NetworkCredential("", "MyMang0Time(9)").SecurePassword);
                ctx.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                {
                    e.WebRequestExecutor.WebRequest.UserAgent = "NONISV|Microsoft|CampusProjectSite/1.0";
                };
                Web web = ctx.Web;
                ctx.Load(web, w => w.Url);
                ctx.Load(web, w => w.Id);
                ctx.ExecuteQueryRetry();
                SiteDesignTemplate.Provision.SiteDesignTemplate.ApplySiteDesign(ctx, new Guid("49b6dc49-27d6-485d-8f82-e70db2c3c523"));
                //SiteDesignTemplate.Provision.SiteDesignTemplate.ApplyPnpTemplate(web);
            }

        }

    }


}
