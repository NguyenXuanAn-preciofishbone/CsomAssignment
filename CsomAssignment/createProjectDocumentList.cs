using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace CsomAssignment
{
    class createProjectDocumentList
    {
        private ClientContext context;
        private Web web;

        public createProjectDocumentList(ClientContext context, SharePointOnlineCredentials credentials)
        {
            this.context = context;
            this.context.Credentials = credentials;
            web = context.Web;
        }

        public void Execute()
        {
            createField();
            createContentType();
            createList();
            addLookupField();
        }

        void createField()
        {
            Web web = context.Web;

            Guid idNTitle = Guid.NewGuid();
            Guid idTypeOfDocument = Guid.NewGuid();

            string fieldNTitle = $"<Field ID='{idNTitle}' DisplayName='NTitle' Name='NTitle' Group='CSOM assignment' Type='Text' />";
            string fieldTypeOfDocument =
                $"<Field ID='{idTypeOfDocument}' DisplayName='Type of document' Name='TypeOfDocment' Group='CSOM assignment' Type='Choice' Format='RadioButtons' Hidden='FALSE'>"
                + "<CHOICES>"
                + "    <CHOICE>Signed</CHOICE>"
                + "    <CHOICE>Design</CHOICE>"
                + "    <CHOICE>Development</CHOICE>"
                + "    <CHOICE>Maintenance</CHOICE>"
                + "    <CHOICE>Closed</CHOICE>"
                + "</CHOICES>"
                + "</Field>";

            web.Fields.AddFieldAsXml(fieldNTitle, false, AddFieldOptions.AddFieldInternalNameHint);
            web.Fields.AddFieldAsXml(fieldTypeOfDocument, false, AddFieldOptions.AddFieldInternalNameHint);

            Console.WriteLine("Success create field");
        }

        void createContentType()
        {
            ContentTypeCollection contentTypes = web.ContentTypes;

            context.Load(contentTypes);
            context.ExecuteQuery();

            ContentType parentContentType = (from contentType in contentTypes where contentType.Name == "Document" select contentType).FirstOrDefault();

            ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation();
            contentTypeCreationInformation.Name = "Project documents";
            contentTypeCreationInformation.Description = "Project documents";
            contentTypeCreationInformation.Group = "CSOM Assignment";
            contentTypeCreationInformation.ParentContentType = parentContentType;

            ContentType newContentType = contentTypes.Add(contentTypeCreationInformation);
            context.Load(newContentType);
            context.ExecuteQuery();

            Field fieldNTitle = web.Fields.GetByInternalNameOrTitle("NTitle");
            Field fieldNDescription = web.Fields.GetByInternalNameOrTitle("NDescription");
            Field fieldTypeOfDocument = web.Fields.GetByInternalNameOrTitle("Type of document");
            ContentType ProjectDocuments = (from c in contentTypes
                                   where c.Name == "Project documents"
                                   select c).FirstOrDefault();

            ProjectDocuments.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldNTitle
            });
            ProjectDocuments.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldNDescription
            });
            ProjectDocuments.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldTypeOfDocument
            });
            ProjectDocuments.Update(true);
            context.ExecuteQuery();

            Console.WriteLine("Success create content type");
        }

        void createList()
        {
            ListCreationInformation listCreationInformation = new ListCreationInformation();
            listCreationInformation.Title = "Project documents";
            listCreationInformation.TemplateType = (int)ListTemplateType.DocumentLibrary;

            List newList = web.Lists.Add(listCreationInformation);
            context.Load(newList);
            context.ExecuteQuery();

            ContentTypeCollection contentTypes = web.ContentTypes;
            context.Load(contentTypes);
            context.ExecuteQuery();

            ContentType ProjectDocumentsContentType = (from contentType in contentTypes where contentType.Name == "Project documents" select contentType).FirstOrDefault();

            List ProjectDocumentsList = web.Lists.GetByTitle("Project documents");
            ProjectDocumentsList.ContentTypes.AddExistingContentType(ProjectDocumentsContentType);
            ProjectDocumentsList.Update();
            context.Web.Update();
            context.ExecuteQuery();

            Console.WriteLine("Success create list");
        }

        void addLookupField()
        {
            List ProjectList = web.Lists.GetByTitle("Project");
            List ProjectDocumentsList = web.Lists.GetByTitle("Project documents");

            context.Load(ProjectList, p => p.Id);
            context.Load(ProjectDocumentsList, pd => pd.Fields);
            context.ExecuteQuery();

            string fieldProjectLinked = @"<Field Type='Lookup' Name='ProjectLinked' StaticName='ProjectLinked' DisplayName='Project linked' List='" + ProjectList.Id + "' ShowField = 'Title'/>";

            ProjectList.Fields.AddFieldAsXml(fieldProjectLinked, true, AddFieldOptions.DefaultValue);

            ProjectList.Update();
            context.ExecuteQuery();

            Console.WriteLine("Success add new lookup field");
        }
    }
}
