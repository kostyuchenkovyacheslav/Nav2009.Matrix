using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Nav2009.Matrix
{
    class XmlDocBuilder
    {
        static XmlDocument xmlDoc = new XmlDocument();

        public void InitXmlDoc() // add parameter - root node name?
        {
            xmlDoc = new XmlDocument();
            XmlNode xmlNode1 = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "yes");
            xmlDoc.AppendChild(xmlNode1);
            xmlNode1 = xmlDoc.CreateElement("TableData");
            xmlDoc.AppendChild(xmlNode1);
        }

        public void xmlDocAddElement(string CaptionAttr) // add parameter - element name? - add attributes?
        {
            XmlNode xmlNode1 = xmlDoc.CreateElement("Record");
            xmlDoc.AppendChild(xmlNode1);
            XmlAttribute xmlAttr1 = xmlDoc.CreateAttribute("Caption");
            xmlAttr1.Value = CaptionAttr;
            xmlNode1.Attributes.Append(xmlAttr1);
        }
    }
}
