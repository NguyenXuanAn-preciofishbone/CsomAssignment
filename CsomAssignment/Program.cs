using System;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Security;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;

namespace CsomAssignment
{
    class Program
    {
        static void Main(string[] args)
        {
            const string username = "lanehacker7294@NguyenXuanAn.onmicrosoft.com";
            const string password = "@Ab0907284582";
            const string adminSite = "https://nguyenxuanan-admin.sharepoint.com/";
            const string rootSite = "https://nguyenxuanan.sharepoint.com/";

            var securedPassword = new SecureString();
            foreach (var c in password.ToCharArray()) securedPassword.AppendChar(c);

            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(username, securedPassword);

            //Console.Write("Input site title: ");
            //string title = Console.ReadLine();
            //string newSite = rootSite + "sites/" + title;

            const string testSite = "https://nguyenxuanan.sharepoint.com/sites/test2";
            using (ClientContext context = new ClientContext(testSite))
            {
                createEmployeeList test = new createEmployeeList(context, credentials);
                test.Execute();
            };
        }
    }
}
