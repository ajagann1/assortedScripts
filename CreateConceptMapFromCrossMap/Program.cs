using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using DocumentFormat.OpenXml.Office.CustomUI;
using Newtonsoft.Json;
using System.Xml;
using DocumentFormat.OpenXml.Math;
using System.Text.Json;
using Hl7.Fhir.Serialization;
using Hl7.Fhir.Model;

class Program
{

    static void Main(String[] args)
    {
        //Cross Map (location is set to a direct filepath right now, modify however)
        XLWorkbook mappings = new XLWorkbook("Input\\Encounter_Type_Mapping_URMC_2023-08-17.xlsx");

        XmlDocument doc = new XmlDocument();
        //Template file (ConceptMap with no mappings)
        doc.Load("Input\\ConceptMapTemplate.xml");
        XmlNode group = doc.GetElementsByTagName("group")[0];

        IXLRow mapRow = mappings.Worksheet("Mappings").Row(2);
        while (!mapRow.IsEmpty())
        {
            //Client codes
            String code = mapRow.Cell("A").Value.ToString().Trim();
            //Measure codes
            String target = mapRow.Cell("C").Value.ToString().Trim();

            if (target != "" && code != "")
            {
                String tar = target.Contains("QM#LOC") ? target.Replace("QM#LOC", "") : target;

                XmlElement c = doc.CreateElement("code");
                c.SetAttribute("value", code);

                XmlElement tCode = doc.CreateElement("code");
                tCode.SetAttribute("value", tar);

                XmlElement tRelationship = doc.CreateElement("relationship");
                tRelationship.SetAttribute("value", "inexact");

                XmlElement tDisplay = doc.CreateElement("display");
                tDisplay.SetAttribute("value", mapRow.Cell("D").Value.ToString().Trim());

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

        var xmlStr = doc.OuterXml;
        var deserializer = new FhirXmlPocoDeserializer(ModelInfo.ModelInspector, new FhirXmlPocoDeserializerSettings
        {
            Validator = null
        });
        Resource conceptMap = null;
        Hl7.Fhir.Model.Base baseRes = null;
        try
        {
            conceptMap = deserializer.DeserializeResource(xmlStr); //Deserialize<ConceptMap>(xmlStr);
        }
        catch (DeserializationFailedException ex)
        {
            baseRes = ex.PartialResult;
        }
        if (conceptMap == null)
            conceptMap = (ConceptMap)baseRes;

        var options = new JsonSerializerOptions().ForFhir(ModelInfo.ModelInspector).Pretty();
        string conceptMapJson = System.Text.Json.JsonSerializer.Serialize(conceptMap, options);

        //Change output name to whatever you'd like
        File.WriteAllText("Output\\CompletedConceptMap.json", conceptMapJson);

        //Change output name to whatever you'd like
        //doc.Save("Output\\CompletedConceptMap.xml");
    }
}