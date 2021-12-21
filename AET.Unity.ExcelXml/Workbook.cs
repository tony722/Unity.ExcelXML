using System;
using System.ComponentModel;
using System.Linq;
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

    public string AETVersion { get; private set; }


      public void LoadXmlData(string xmlData) {
      xDoc = XDocument.Parse(xmlData);
      xRoot = xDoc.Root;
      Settings.ssPrefix = xRoot.GetNamespaceOfPrefix("ss");
      Worksheets = new Worksheets(xRoot.Elements(Settings.ssPrefix + "Worksheet"));
      ReadAETVersion();
      }

      private void ReadAETVersion() {
        var name = xRoot.Element(Settings.ssPrefix + "Names");
        if (name == null) return;
        var names = name.Elements();
        var aetVersion = names.FirstOrDefault( x => x.Attribute(Settings.ssPrefix + "Name").Value == "AETVersion" || x.Attribute(Settings.ssPrefix + "Name").Value == "AET_Version");
        if(aetVersion == null) return;
        var refersTo = aetVersion.Attribute(Settings.ssPrefix + "RefersTo");
        var value = refersTo.Value.Substring(1);
        value = value.Replace("\"","");
        AETVersion = value;
      }
  }
}


