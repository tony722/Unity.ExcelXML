using Crestron.SimplSharp.CrestronXmlLinq;

namespace AET.Unity.ExcelXml {
  public class Worksheet {
    public Worksheet() { }
    public Worksheet(XElement worksheetElement) {
      XElement = worksheetElement;
      Rows = new Rows(worksheetElement.Element(Settings.ssPrefix + "Table"));
      Columns = new Columns(Rows);
    }
    internal XElement XElement { get; set; }

    public string Name { get { return XElement.Attribute((XName)(Settings.ssPrefix + "Name")).Value; }  }

    public Rows Rows { get; private set; }

    public Columns Columns { get; private set; }
  }
}