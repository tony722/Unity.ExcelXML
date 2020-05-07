using System.Collections.Generic;

namespace Unity.ExcelXml {
  public class Columns : IEnumerable<Column> {
    private List<Column> columns;
    public Columns(Rows rows) {
      columns = new List<Column>();
      for (int i = 0; i < rows.ColumnsCount; i++) {
        var column = new Column(rows, i);
        columns.Add(column);
      }
    }

    public Column this[int index] { get { return columns[index]; } }

    public int Count { get { return columns.Count; } }

    #region IEnumerable<Column> Members

    IEnumerator<Column> IEnumerable<Column>.GetEnumerator() {
      return columns.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
      return columns.GetEnumerator();
    }

    #endregion
  }
}