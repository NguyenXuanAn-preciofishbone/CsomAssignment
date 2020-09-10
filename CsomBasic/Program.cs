using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace CsomBasic
{
    class Program
    {
        static string siteUrl = "https://nguyenxuanan.sharepoint.com/sites/TrainingAssignment/HR_department";
        static string loginName = "lanehacker7294@nguyenxuanan.onmicrosoft.com";
        static string password = "@Ab0907284582";
        static SecureString securePassword = new SecureString();

        static void Main(string[] args)
        {
            password.ToCharArray().ToList().ForEach(c => securePassword.AppendChar(c));
            createFieldByXml();
            //createFieldByObject();
            //createContentType();
            //createList();
        }

        static void createFieldByXml()
        {
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(loginName, securePassword);

                var web = context.Web;
                web.Fields.AddFieldAsXml("<Field DisplayName='Choice test' Name='ChoiceTestDif' Group='SharePoint Saturday 2014 Columns' Type='Choice' />", false, AddFieldOptions.AddFieldInternalNameHint);

                context.ExecuteQuery();

                Console.WriteLine("success create new field");
            };
        }

        static void createFieldByObject()
        {
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(loginName, securePassword);

                var fieldInfo = new FieldCreationInformation();
                fieldInfo.FieldType = FieldType.Geolocation;
                fieldInfo.InternalName = "test";
                fieldInfo.DisplayName = "test";
                context.Site.RootWeb.Fields.Add(fieldInfo);
                context.ExecuteQuery();
            };
        }

        static void createContentType()
        {
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(loginName, securePassword);

                ContentTypeCollection oContentTypeCollection = context.Site.RootWeb.ContentTypes;

                //Load content type collection
                context.Load(oContentTypeCollection);
                context.ExecuteQuery();

                //Give parent content type name over here
                ContentType oparentContentType = (from contentType in oContentTypeCollection where contentType.Name == "Document" select contentType).First();

                ContentTypeCreationInformation oContentTypeCreationInformation = new ContentTypeCreationInformation();

                //Name of the new content type
                oContentTypeCreationInformation.Name = "New Document Content Type";

                //Description of the new content type
                oContentTypeCreationInformation.Description = "New Document Content Type Description";

                //Name of the group under which the new content type will be creted
                oContentTypeCreationInformation.Group = "Custom Content Types Group";

                //Specify the parent content type over here
                oContentTypeCreationInformation.ParentContentType = oparentContentType;

                //Add "ContentTypeCreationInformation" object created above
                ContentType oContentType = oContentTypeCollection.Add(oContentTypeCreationInformation);

                context.ExecuteQuery();
            };
        }

        static void createList()
        {
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(loginName, securePassword);

                // The properties of the new custom list
                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = "ListTitle";
                creationInfo.Description = "New list description";
                creationInfo.TemplateType = (int)ListTemplateType.GenericList;

                List newList = context.Web.Lists.Add(creationInfo);
                context.Load(newList);
                // Execute the query to the server.
                context.ExecuteQuery();

                List olist = context.Web.Lists.GetByTitle("ListTitle");
                olist.ContentTypesEnabled = true;
                olist.Update();
                context.ExecuteQuery();

                ContentTypeCollection contentTypeCollection;

                // Option - 1 - Get Content Types from Root web
                contentTypeCollection = context.Site.RootWeb.ContentTypes;

                context.Load(contentTypeCollection);
                context.ExecuteQuery();

                // Get the content type from content type collection. Give the content type name over here
                ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "New Document Content Type" select contentType).First();

                // Add existing content type on target list. Give target list name over here.
                List targetList = context.Web.Lists.GetByTitle("ListTitle");
                targetList.ContentTypes.AddExistingContentType(targetContentType);
                targetList.Update();
                context.Web.Update();
                context.ExecuteQuery();
            };
        }
    }
}
