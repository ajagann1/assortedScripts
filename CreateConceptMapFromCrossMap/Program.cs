using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using DocumentFormat.OpenXml.Office.CustomUI;
using Newtonsoft.Json;
using System.Xml;

class Program
{

    static void Main(String[] args)
    {
        //Cross Map (location is set to a direct filepath right now, modify however)
        XLWorkbook mappings = new XLWorkbook("Input\\Yale New Haven Location Mappings.xlsx");

        XmlDocument doc = new XmlDocument();
        //Template file (ConceptMap with no mappings)
        doc.Load("Input\\ConceptMapTemplate.xml");
        XmlNode group = doc.GetElementsByTagName("group")[0];

        IXLRow mapRow = mappings.Worksheet(1).Row(2);
        while (!mapRow.IsEmpty())
        {
            //Client codes
            String code = mapRow.Cell("D").Value.ToString().Trim();
            //Measure codes
            String target = mapRow.Cell("Q").Value.ToString().Trim();

            if (target != "")
            {
                String tar = target.Contains("QM#LOC") ? target.Replace("QM#LOC", "") : target;

                XmlElement c = doc.CreateElement("code");
                c.SetAttribute("value", code);

                XmlElement tCode = doc.CreateElement("code");
                tCode.SetAttribute("value", tar);

                XmlElement tRelationship = doc.CreateElement("relationship");
                tRelationship.SetAttribute("value", "inexact");

                XmlElement tDisplay = doc.CreateElement("display");
                tDisplay.SetAttribute("value", mapRow.Cell("P").Value.ToString().Trim());

                XmlElement t = doc.CreateElement("target");
                t.AppendChild(tCode);
                t.AppendChild(tRelationship);
                t.AppendChild(tDisplay);

                XmlElement element = doc.CreateElement("element");
                element.AppendChild(c);
                element.AppendChild(t);

                group.AppendChild(element);
            }

            mapRow = mapRow.RowBelow();
        }
        //Change output name to whatever you'd like
        doc.Save("Output\\CompletedConceptMap.xml");
    }
}