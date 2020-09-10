using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsomAssignment
{
    class createProjectList
    {
        private ClientContext context;
        private Web web;
        private Guid idProjectName;
        private Guid idNDescription;
        private Guid idState;
        public createProjectList(ClientContext context, SharePointOnlineCredentials credentials)
        {
            this.context = context;
            this.context.Credentials = credentials;
            web = this.context.Web;
            idProjectName = Guid.NewGuid();
            idNDescription = Guid.NewGuid();
            idState = Guid.NewGuid();
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
            string fieldProjectName = $"<Field ID='{idProjectName}' DisplayName='Project name' Name='ProjectName' Group='CSOM assignment' Type='Text' />";
            string fieldNDescription = $"<Field ID='{idNDescription}' DisplayName='NDescription' Name='NDescription' Group='CSOM assignment' Type='Text' />";
            string fieldState =
                $"<Field ID='{idState}' DisplayName='State' Name='State' Group='CSOM assignment' Type='Choice' Format='RadioButtons' Hidden='FALSE'>"
                + "<CHOICES>"
                + "    <CHOICE>Signed</CHOICE>"
                + "    <CHOICE>Design</CHOICE>"
                + "    <CHOICE>Development</CHOICE>"
                + "    <CHOICE>Maintenance</CHOICE>"
                + "    <CHOICE>Closed</CHOICE>"
                + "</CHOICES>"
                + "</Field>";

            web.Fields.AddFieldAsXml(fieldProjectName, false, AddFieldOptions.AddFieldInternalNameHint);
            web.Fields.AddFieldAsXml(fieldNDescription, false, AddFieldOptions.AddFieldInternalNameHint);
            web.Fields.AddFieldAsXml(fieldState, false, AddFieldOptions.AddFieldInternalNameHint);

            context.ExecuteQuery();

            Console.WriteLine("Success create field");
        }

        void createContentType()
        {
            ContentTypeCollection contentTypes = web.ContentTypes;

            context.Load(contentTypes);
            context.ExecuteQuery();

            ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation();
            contentTypeCreationInformation.Name = "Project";
            contentTypeCreationInformation.Description = "Project";
            contentTypeCreationInformation.Group = "CSOM Assignment";

            ContentType newContentType = contentTypes.Add(contentTypeCreationInformation);
            context.Load(newContentType);
            context.ExecuteQuery();

            Field fieldProjectName = web.Fields.GetByInternalNameOrTitle("Project name");
            Field fieldStartDate = web.Fields.GetByInternalNameOrTitle("Start Date");
            Field fieldEndDate = web.Fields.GetByInternalNameOrTitle("End Date");
            Field fieldNDescription = web.Fields.GetByInternalNameOrTitle("NDescription");
            Field fieldState = web.Fields.GetByInternalNameOrTitle("State");

            ContentType Project = (from c in contentTypes
                                   where c.Name == "Project"
                                   select c).FirstOrDefault();

            Project.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldProjectName
            });
            Project.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldStartDate
            });
            Project.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldEndDate
            });
            Project.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldNDescription
            });
            Project.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = fieldState
            });
            Project.Update(true);
            context.ExecuteQuery();

            Console.WriteLine("Success create content type");
        }

        void createList()
        {
            ListCreationInformation listCreationInformation = new ListCreationInformation();
            listCreationInformation.Title = "Project";
            listCreationInformation.TemplateType = (int)ListTemplateType.GenericList;

            List newList = web.Lists.Add(listCreationInformation);
            context.Load(newList);
            context.ExecuteQuery();

            ContentTypeCollection contentTypes = web.ContentTypes;
            context.Load(contentTypes);
            context.ExecuteQuery();

            ContentType ProjectContentType = (from contentType in contentTypes where contentType.Name == "Project" select contentType).FirstOrDefault();

            List ProjectList = web.Lists.GetByTitle("Project");
            ProjectList.ContentTypes.AddExistingContentType(ProjectContentType);
            ProjectList.Update();
            context.Web.Update();
            context.ExecuteQuery();

            Console.WriteLine("Success create list");
        }

        void addLookupField()
        {
            List ProjectList = web.Lists.GetByTitle("Project");
            List EmployeeList = web.Lists.GetByTitle("Employee");

            context.Load(ProjectList, p => p.Fields);
            context.Load(EmployeeList, e => e.Id);
            context.ExecuteQuery();

            string fieldLeader = @"<Field Type='Lookup' Name='Leader' StaticName='Leader' DisplayName='Leader' List='" + EmployeeList.Id + "' ShowField = 'Title' />";
            string fieldMembers = @"<Field Type='LookupMulti' Name='Members' StaticName='Members' DisplayName='Members' List='" + EmployeeList.Id + "' ShowField = 'Title' Mult = 'TRUE' />";

            ProjectList.Fields.AddFieldAsXml(fieldLeader, true, AddFieldOptions.DefaultValue);
            ProjectList.Fields.AddFieldAsXml(fieldMembers, true, AddFieldOptions.DefaultValue);

            ProjectList.Update();
            context.ExecuteQuery();

            Console.WriteLine("Success add new lookup field");
        }
    }
}
