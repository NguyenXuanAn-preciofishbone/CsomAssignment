﻿using System;
using System.Security;
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
            const string testSite = "https://nguyenxuanan.sharepoint.com/sites/example";
            const string rootSite = "https://nguyenxuanan.sharepoint.com/";

            var securedPassword = new SecureString();
            foreach (var c in password.ToCharArray()) securedPassword.AppendChar(c);

            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(username, securedPassword);

            int flag = 0;
            while (flag == 0)
            {
                Console.WriteLine("Choose your option: ");
                Console.WriteLine("1. Create Employee list ");
                Console.WriteLine("2. Create Project list");
                Console.WriteLine("3. Create Project documents list ");
                Console.WriteLine("4. Create new Site with all of the above list");
                Console.WriteLine("WARNING: You must choose option 1 before 2, 2 before 3");
                Console.WriteLine("Input any other key to exist");
                Console.Write("Your choice: ");
                string input = Console.ReadLine();
                switch (input)
                {
                    case "1":
                        using (ClientContext context = new ClientContext(testSite))
                        {
                            createEmployeeList operation = new createEmployeeList(context, credentials);
                            operation.Execute();
                        }
                        break;
                    case "2":
                        using (ClientContext context = new ClientContext(testSite))
                        {
                            createProjectList operation = new createProjectList(context, credentials);
                            operation.Execute();
                        }
                        break;
                    case "3":
                        using (ClientContext context = new ClientContext(testSite))
                        {
                            createProjectDocumentList operation = new createProjectDocumentList(context, credentials);
                            operation.Execute();
                        }
                        break;
                    case "4":
                        using (ClientContext context = new ClientContext(adminSite))
                        {
                            createFullSite operation = new createFullSite(context, credentials);
                            operation.Execute(rootSite, username);
                        }
                        break;
                    default:
                        flag = 1;
                        break;
                }
            }
        }
    }
}
