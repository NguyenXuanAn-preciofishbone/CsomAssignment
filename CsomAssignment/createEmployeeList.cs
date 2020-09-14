using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace CsomAssignment
{
    class createEmployeeList
    {
        private ClientContext context;
        private Web web;
        private Guid idLastName;
        private Guid idShortDescription;
        private Guid idProgrammingLanguage;

        public createEmployeeList(ClientContext context, SharePointOnlineCredentials credentials)
        {
            this.context = context;
            this.context.Credentials = credentials;
            web = context.Web;
            idLastName = Guid.NewGuid();
            idShortDescription = Guid.NewGuid();
            idProgrammingLanguage = Guid.NewGuid();
        }

        public void Execute()
        {
            createField();
            createContentType();
            createList();
        }

        void createField()
        {
            string fieldLastName = $"<Field ID='{idLastName}' DisplayName='Last name' Name='LastName' Group='CSOM assignment' Type='Text' />";
            string fieldShortDescription = $"<Field ID='{idShortDescription}' DisplayName='Short description' Name='ShortDescription' Group='CSOM assignment' Type='HTML' />";
            string fieldProgrammingLanguage =
                $"<Field ID='{idProgrammingLanguage}' DisplayName='Programming language' Name='ProgrammingLanguage' Group='CSOM assignment' Type='Choice' Format='Dropdown' Hidden='FALSE'>"
                + "<CHOICES>"
                + "    <CHOICE>C#</CHOICE>"
                + "    <CHOICE>F#</CHOICE>"
                + "    <CHOICE>Visual Basic</CHOICE>"
                + "    <CHOICE>Java</CHOICE>"
                + "    <CHOICE>JQuery</CHOICE>"
                + "    <CHOICE>AngularJS</CHOICE>"
                + "    <CHOICE>Other</CHOICE>"
                + "</CHOICES>"
                + "</Field>";

            web.Fields.AddFieldAsXml(fieldLastName, false, AddFieldOptions.AddFieldInternalNameHint);
            web.Fields.AddFieldAsXml(fieldShortDescription, false, AddFieldOptions.AddFieldInternalNameHint);
            web.Fields.AddFieldAsXml(fieldProgrammingLanguage, false, AddFieldOptions.AddFieldInternalNameHint);

            context.ExecuteQuery();

            Console.WriteLine("Success create field");
        }

        void createContentType()
        {
            ContentTypeCollection contentTypes = web.ContentTypes;

            context.Load(contentTypes);
            context.ExecuteQuery();

            ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation();
            contentTypeCreationInformation.Name = "Employee";
            contentTypeCreationInformation.Description = "Employee";
            contentTypeCreationInformation.Group = "CSOM Assignment";

            ContentType newContentType = contentTypes.Add(contentTypeCreationInformation);
            context.Load(newContentType);
            context.ExecuteQuery();

            Field fieldFirstName = web.Fields.GetByInternalNameOrTitle("First Name");
            Field fieldLastName = web.Fields.GetByInternalNameOrTitle("Last name");
            Field fieldEmail = web.Fields.GetByInternalNameOrTitle("E-Mail");
            Field fieldShortDescription = web.Fields.GetByInternalNameOrTitle("Short description");
            Field fieldProgrammingLanguage = web.Fields.GetByInternalNameOrTitle("Programming language");

            ContentType Employee = (from c in contentTypes
                                    where c.Name == "Employee"
                                    select c).FirstOrDefault();

            Employee.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldFirstName
            });
            Employee.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldLastName
            });
            Employee.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldEmail
            });
            Employee.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldShortDescription
            });
            Employee.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldProgrammingLanguage
            });
            Employee.Update(true);
            context.ExecuteQuery();

            Console.WriteLine("Success create content type");
        }

        void createList()
        {
            ListCreationInformation listCreationInformation = new ListCreationInformation();
            listCreationInformation.Title = "Employee";
            listCreationInformation.TemplateType = (int)ListTemplateType.GenericList;

            List newList = web.Lists.Add(listCreationInformation);
            context.Load(newList);
            context.ExecuteQuery();

            ContentTypeCollection contentTypes = web.ContentTypes;
            context.Load(contentTypes);
            context.ExecuteQuery();

            ContentType EmployeeContentType = (from contentType in contentTypes where contentType.Name == "Employee" select contentType).FirstOrDefault();

            List EmployeeList = web.Lists.GetByTitle("Employee");
            EmployeeList.ContentTypes.AddExistingContentType(EmployeeContentType);
            EmployeeList.Update();
            context.Web.Update();
            context.ExecuteQuery();

            Console.WriteLine("Success create list");
        }
    }
}
