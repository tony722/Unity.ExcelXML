namespace Unity.ExcelXml {
  public class Column {    
    public Column(Rows rows, int columnIndex) {
      this.Cells = new Cells(rows, columnIndex);
    }
    
    public string this[int index] {
      get { return Cells[index]; }
    }

    public Cells Cells { get; private set; }    
    
  }
}