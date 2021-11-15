using Crestron.SimplSharp.CrestronXmlLinq;

namespace AET.Unity.ExcelXml {
  public class Row {
    public Row(XElement rowXml, int columnsCount, Rows parentRows) {
      RowXML = rowXml;
      ColumnsCount = columnsCount;
      Cells = new Cells(this);
      ParentRows = parentRows;
    }

    public Rows ParentRows { get; private set; }

    public XElement RowXML { get; private set; }

    public int ColumnsCount { get; private set; }

    public string this[int index] {
      get { return Cells[index]; }
    }

    public string this[string columnName] {
      get { return Cells[columnName]; }
    }

    public Cells Cells { get; private set; }

    public Row Clone() {
      var rowXMLcopy = new XElement(RowXML);
      var rowCopy = new Row(rowXMLcopy, ColumnsCount, ParentRows);
      return rowCopy;
    }

    //public string[] GetCells() { return cells.Select(c => c.Value).ToArray(); }
    //public ReadOnlyCollection<Cell> Cells { get { return cells.AsReadOnly(); } }
    //public Cell Cell(string name) {
    //  return cells.FirstOrDefault(c => c.Name == name);
    //}

  }
}