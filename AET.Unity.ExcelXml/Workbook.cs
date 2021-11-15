using Crestron.SimplSharp.CrestronXmlLinq;

namespace AET.Unity.ExcelXml {
  public class Workbook {    
    public Workbook() { }
   
    public XElement xRoot { get; private set; }
    public XDocument xDoc { get; private set; }
    public Worksheets Worksheets { get; private set; }
    public Worksheet this[string name] {
      get { return Worksheets[name]; }
    }

    public void LoadXmlData(string xmlData) {
      xDoc = XDocument.Parse(xmlData);
      xRoot = xDoc.Root;
      Settings.ssPrefix = xRoot.GetNamespaceOfPrefix("ss");
      Worksheets = new Worksheets(xRoot.Elements(Settings.ssPrefix + "Worksheet"));
    }
  }
}


