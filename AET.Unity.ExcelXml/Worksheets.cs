using System.Collections.Generic;
using System.Linq;
using Crestron.SimplSharp.CrestronXmlLinq;

namespace AET.Unity.ExcelXml {
  public class Worksheets : IEnumerable<Worksheet> {
    private List<Worksheet> worksheets;
    public Worksheets(IEnumerable<XElement> worksheetsXml) {
      this.worksheets = worksheetsXml.Select(x => new Worksheet(x)).ToList();
    }

    public Worksheet this[int index] {
      get { return worksheets[index]; }
    }

    public Worksheet this[string name] {
      get { return worksheets.Single(ws => ws.Name == name); }
    }

    public bool Contains(string name) {
      return worksheets.Any(ws => ws.Name == name);
    }

    #region IEnumerable<Worksheet> Members

    IEnumerator<Worksheet> IEnumerable<Worksheet>.GetEnumerator() {
      return worksheets.GetEnumerator();
    }

    #endregion

    #region IEnumerable Members

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
      return worksheets.GetEnumerator();
    }

    #endregion
  }
}