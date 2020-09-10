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
        private Guid idDescription;
        private Guid idState;
        public createProjectList(ClientContext context, SharePointOnlineCredentials credentials)
        {
            this.context = context;
            this.context.Credentials = credentials;
            web = this.context.Web;
            idProjectName = Guid.NewGuid();
            idDescription = Guid.NewGuid();
            idState = Guid.NewGuid();
        }
        public void Execute()
        {
            createField();
            createContentType();
            addLookupField();
            createList();
        }

        void createField()
        {
            try
            {
                string fieldProjectName = $"<Field ID='{idProjectName}' DisplayName='Project name' Name='ProjectName' Group='CSOM assignment' Type='HTML' />";
                string fieldDescription = $"<Field ID='{idDescription}' DisplayName='Description' Name='Description' Group='CSOM assignment' Type='HTML' />";
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
                web.Fields.AddFieldAsXml(fieldDescription, false, AddFieldOptions.AddFieldInternalNameHint);
                web.Fields.AddFieldAsXml(fieldState, false, AddFieldOptions.AddFieldInternalNameHint);
                            catch (Microsoft.SharePoint.Client.ServerException e)
            {
                Console.WriteLine("Field already created. Skip creating new field");
                return;
            }
            Console.WriteLine("Success create new field");
        }

        void createContentType()
        {

        }

        void addLookupField()
        {

        }

        void createList()
        {

        }
    }
}
