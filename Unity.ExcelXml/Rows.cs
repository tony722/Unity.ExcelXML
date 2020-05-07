using System;
using System.Collections.Generic;
using System.Linq;
using Crestron.SimplSharp.CrestronXmlLinq;

namespace Unity.ExcelXml {
  public class Rows : IEnumerable<Row> {
    private List<Row> rows;
    private XElement table;
    public Rows(XElement table) {
      this.table = table;
      ColumnsCount = int.Parse(table.Attribute(Settings.ssPrefix + "ExpandedColumnCount").Value);
      var rowsXml = table.Descendants(Settings.ssPrefix + "Row");
      this.rows = rowsXml.Select(
        r => {
          var row = new Row(r, ColumnsCount, this);
          return row; }).ToList();
    }
    internal Rows() { }
    public Dictionary<string, int> ColumnNames { get; private set; }

    public Row Add() {
      IncrementXmlTableRowCount();
      var newRowXElement = new XElement(Settings.ssPrefix + "Row", new XAttribute(Settings.ssPrefix + "AutoFitHeight", "0"));
      table.Add(newRowXElement);
      var row = new Row(newRowXElement, ColumnsCount, this);
      rows.Add(row);
      return row;
    }

    public void SetColumnNames(IEnumerable<string> columns) {
      ColumnNames = columns.Select((value, index) => new { value, index })
        .ToDictionary(o => o.value, o => o.index, StringComparer.InvariantCultureIgnoreCase);
    }

    public void SetColumnNamesFromRow(int rowIndex) {
      var row = rows[rowIndex];
      SetColumnNames(row.Cells.Select(c => c.Value));
    }

    private void IncrementXmlTableRowCount() {
      if (!table.HasAttribute(Settings.ssPrefix + "ExpandedRowCount")) return;
      var rowCount = int.Parse(table.Attribute(Settings.ssPrefix + "ExpandedRowCount").Value);
      table.Attribute(Settings.ssPrefix + "ExpandedRowCount").Value = (rowCount + 1).ToString();
    }

    public void Remove(Row row) {
      row.RowXML.Remove();
      rows.Remove(row);
    }

    public int ColumnsCount { get; private set; }
    public int Count { get { return rows.Count; } }

    public Row this[int index] { 
      get { return rows[index]; }
    }

    #region IEnumerable Members

    IEnumerator<Row> IEnumerable<Row>.GetEnumerator() {
      return rows.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
      return rows.GetEnumerator();
    }

    #endregion
  }
}