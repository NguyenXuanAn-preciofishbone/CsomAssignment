using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsomAssignment
{
    class createProjectDocumentList
    {
        private ClientContext context;
        createProjectDocumentList(ClientContext context, SharePointOnlineCredentials credentials)
        {
            this.context = context;
            this.context.Credentials = credentials;
        }

        void createField()
        {
            Web web = context.Web;

            Guid idTitle = Guid.NewGuid();
            Guid idTypeOfDocument = Guid.NewGuid();

            string fieldTitle = $"<Field ID='{idTitle}' DisplayName='Last name' Name='LastName' Group='CSOM assignment' Type='Text' />";
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

            web.Fields.AddFieldAsXml(fieldTitle, false, AddFieldOptions.AddFieldInternalNameHint);
            web.Fields.AddFieldAsXml(fieldTypeOfDocument, false, AddFieldOptions.AddFieldInternalNameHint);
        }

        void createContentType()
        {

        }

        void createList()
        {

        }
    }
}
