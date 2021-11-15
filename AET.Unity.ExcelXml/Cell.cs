using Crestron.SimplSharp.CrestronXmlLinq;

namespace AET.Unity.ExcelXml {
  public class Cell {
    private readonly Row parentRow;

    public Cell(Row parentRow) {
      this.parentRow = parentRow;
      CellXml = new XElement(Settings.ssPrefix + "Cell", new XElement(Settings.ssPrefix + "Data", "", new XAttribute(Settings.ssPrefix + "Type", "String")));
    }

    public Cell(XElement cellXml, Row parentRow) {
      CellXml = cellXml;
      this.parentRow = parentRow;
      if (!cellXml.HasElement(Settings.ssPrefix + "Data")) {
        cellXml.Add(new XElement(Settings.ssPrefix + "Data"));
      }
    }

    public XElement CellXml { get; private set; }

    public static implicit operator string(Cell cell) { return cell.Value; }

    public string Value {
      get { return CellXml.Element(Settings.ssPrefix + "Data").Value; }
      set {
        CellXml.Element(Settings.ssPrefix + "Data").Value = value;
        if (IsFillerCell) {
          AddFillerCellToDOM(this);
        }
      }
    }

    public override string ToString() {
      return Value;
    }

    public bool IsFillerCell {
      get { return CellXml.HasAttribute("Filler"); }
      set {
        if (value) {
          if(!IsFillerCell) CellXml.Add(new XAttribute("Filler", true));
        } else {
          if(IsFillerCell) CellXml.Attribute("Filler").Remove();
        }
      }
    }

    public bool IsIndexed {
      get { return CellXml.HasAttribute(Settings.ssPrefix + "Index"); }
      set {
        if (value) {
          if(!IsIndexed) CellXml.Add(new XAttribute(Settings.ssPrefix + "Index", true));
        } else {
          if(IsIndexed) CellXml.Attribute(Settings.ssPrefix + "Index").Remove();
        }
      }
    }

    public int Index {
      get { return IsIndexed ? int.Parse(CellXml.Attribute(Settings.ssPrefix + "Index").Value) : 0; }
      set {
        if (!IsIndexed) CellXml.Add(new XAttribute(Settings.ssPrefix + "Index", value));
        else CellXml.Attribute(Settings.ssPrefix + "Index").Value = value.ToString();
      }
    }

    private void AddFillerCellToDOM(Cell cell) {
      cell.IsFillerCell = false;
      var prev = GetPrevCell(cell);
      if (prev == null) parentRow.RowXML.AddFirst(cell.CellXml);
      else prev.CellXml.AddAfterSelf(cell.CellXml);
    }

    private Cell GetPrevCell(Cell cell) {
      var prev = parentRow.Cells.Previous(cell);
      if (prev.IsFillerCell == false) return prev;
      cell.Index = parentRow.Cells.IndexOf(cell) + 1;
      while (prev.IsFillerCell) {
        prev = parentRow.Cells.Previous(prev);
        if (prev == null) return null;
      }
      return prev;
    }
  }
}