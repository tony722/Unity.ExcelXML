using System;
using System.Collections.Generic;
using System.Linq;

namespace AET.Unity.ExcelXml {
  public class Cells : IEnumerable<Cell> {
    private List<Cell> cells;
    private readonly Row row;

    public Cells(Row row) {
      this.row = row;
      FillCellsList();
    }

    public Cells(Rows rows, int columnIndex) {
      cells = rows.Select(rw => rw.Cells[columnIndex]).ToList();
    }

    public Cell this[int index] { 
      get { return cells[index]; }
      set { cells[index] = value; }
    }

    public Cell this[string columnName] {
      get {
        if(row.ParentRows == null || row.ParentRows.ColumnNames == null) throw new NullReferenceException(string.Format("Column names have not been set. Cannot lookup column by name \"{0}\".", columnName));
        int index;
        if(row.ParentRows.ColumnNames.TryGetValue(columnName, out index)) return cells[index];
        return new Cell(row);
      }
      set { cells[row.ParentRows.ColumnNames[columnName]] = value; }
    }

    public int Count { get { return cells.Count; } } 

    private void FillCellsList() {
      var cellsXml = row.RowXML.Descendants(Settings.ssPrefix + "Cell");
      cells = new List<Cell>();
      var i = 0;
      foreach (var x in cellsXml) {
        var cell = new Cell(x, row);
        if (cell.IsIndexed) {
          AddFillerCells(ref i, cell.Index);
        }
        cells.Add(new Cell(x, row));
        i++;
      }
      AddFillerCells(ref i, row.ColumnsCount + 1);      
    }

    private void AddFillerCells(ref int i, int index) {
      while (i < index - 1) {
        var cell = new Cell(row) { IsFillerCell = true };
        cells.Add(cell);
        i++;
      }
    }

    public Cell Previous(Cell cell) {
      return cells.Previous(cell);
    }

    public int IndexOf(Cell cell) {
      return cells.IndexOf(cell);
    }

    #region IEnumerable Members

    IEnumerator<Cell> IEnumerable<Cell>.GetEnumerator() {
      return cells.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
      return cells.GetEnumerator();
    }

    #endregion
  }
}