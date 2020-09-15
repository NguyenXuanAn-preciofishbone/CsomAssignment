using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;

namespace CsomAssignment
{
    class createFullSite
    {
        private ClientContext context;
        private createEmployeeList employeeList;
        private createProjectList projectList;
        private createProjectDocumentList projectDocumentList;
        private SharePointOnlineCredentials credentials;

        public createFullSite(ClientContext context, SharePointOnlineCredentials credentials)
        {
            this.context = context;
            this.credentials = credentials;
            context.Credentials = this.credentials;
        }

        public void Execute(string rootSite, string username)
        {
            Console.WriteLine("Input site title: ");
            string title = Console.ReadLine();
            string fullUrl = rootSite + "sites/" + title;

            createSite(fullUrl, username, title);
            createComponent(fullUrl);

        }

        public void createSite(String url, String owner, String title = null, String template = "STS#0", uint? localeId = null, int? compatibilityLevel = null, long? storageQuota = null, double? resourceQuota = null, int? timeZoneId = null)
        {
            var tenant = new Tenant(context);

            if (url == null)
                throw new ArgumentException("Site Url must be specified");

            if (string.IsNullOrEmpty(owner))
                throw new ArgumentException("Site Owner must be specified");

            var siteCreationProperties = new SiteCreationProperties { Url = url, Owner = owner };
            if (!string.IsNullOrEmpty(template))
                siteCreationProperties.Template = template;
            if (!string.IsNullOrEmpty(title))
                siteCreationProperties.Title = title;
            if (localeId.HasValue)
                siteCreationProperties.Lcid = localeId.Value;
            if (compatibilityLevel.HasValue)
                siteCreationProperties.CompatibilityLevel = compatibilityLevel.Value;
            if (storageQuota.HasValue)
                siteCreationProperties.StorageMaximumLevel = storageQuota.Value;
            if (resourceQuota.HasValue)
                siteCreationProperties.UserCodeMaximumLevel = resourceQuota.Value;
            if (timeZoneId.HasValue)
                siteCreationProperties.TimeZoneId = timeZoneId.Value;
            var siteOp = tenant.CreateSite(siteCreationProperties);
            context.Load(siteOp);
            context.ExecuteQuery();          

            context.Load(siteOp, i => i.IsComplete);
            context.ExecuteQuery();

            while (!siteOp.IsComplete)
            {
                Console.WriteLine("Creating");
                System.Threading.Thread.Sleep(20000);
                siteOp.RefreshLoad();
                context.ExecuteQuery();
            }

            Console.WriteLine("SiteCollection Created.");
        }

        public void createComponent(string fullUrl)
        {
            ClientContext newSiteContext = new ClientContext(fullUrl);

            employeeList = new createEmployeeList(newSiteContext, this.credentials);
            employeeList.Execute();

            projectList = new createProjectList(newSiteContext, this.credentials);
            projectList.Execute();

            projectDocumentList = new createProjectDocumentList(newSiteContext, this.credentials);
            projectDocumentList.Execute();
        }
    }
}
