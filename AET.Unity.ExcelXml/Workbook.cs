using Crestron.SimplSharp.CrestronXmlLinq;

namespace AET.Unity.ExcelXml {
  public class Workbook {    
    public Workbook() { }
   
    public XElement xmlDoc { get; private set; }

    public Worksheets Worksheets { get; private set; }
    public Worksheet this[string name] {
      get { return Worksheets[name]; }
    }

    public void LoadXmlData(string xmlData) {
      var xml = XDocument.Parse(xmlData);
      xmlDoc = xml.Root;
      Settings.ssPrefix = xmlDoc.GetNamespaceOfPrefix("ss");
      this.Worksheets = new Worksheets(xmlDoc.Elements(Settings.ssPrefix + "Worksheet"));
    }
  }
}


