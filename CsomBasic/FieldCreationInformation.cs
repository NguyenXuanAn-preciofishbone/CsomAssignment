using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace CsomBasic
{
    [XmlRoot("Field")]
    public class FieldCreationInformation
    {
        [XmlAttribute("ID")]
        public Guid Id { get; set; }

        [XmlAttribute()]
        public string DisplayName { get; set; }

        [XmlAttribute("Name")]
        public string InternalName { get; set; }

        [XmlIgnore()]
        public bool AddToDefaultView { get; set; }


        //public IEnumerable<KeyValuePair<string, string>> AdditionalAttributes { get; set; }

        [XmlAttribute("Type")]
        public FieldType FieldType { get; set; }

        [XmlAttribute()]
        public string Group { get; set; }

        [XmlAttribute()]
        public bool Required { get; set; }


        public string ToXml()
        {
            var serializer = new XmlSerializer(GetType());
            var settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.OmitXmlDeclaration = true;
            var emptyNamepsaces = new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });

            using (var stream = new StringWriter())
            using (var writer = XmlWriter.Create(stream, settings))
            {
                serializer.Serialize(writer, this, emptyNamepsaces);
                return stream.ToString();
            }
        }



        public FieldCreationInformation()
        {
            Id = Guid.NewGuid();
        }

    }

    public static class FieldCollectionExtensions
    {
        public static Field Add(this FieldCollection fields, FieldCreationInformation info)
        {
            var fieldSchema = info.ToXml();
            return fields.AddFieldAsXml(fieldSchema, info.AddToDefaultView, AddFieldOptions.AddFieldToDefaultView);
        }
    }
}
