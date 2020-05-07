using System.Linq;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Unity.ExcelXml;

namespace Unity.ExcelXML.Tests {
  [TestClass]
  public class WorksheetTests {
    private Workbook workbook;
    private Worksheet events;

    [TestInitialize]
    public void TestInit() {
      var xmlData = System.IO.File.ReadAllText("Schedule.xml");
      workbook = new Workbook();
      workbook.LoadXmlData(xmlData);
      events = workbook["Events"];
    }

    [TestMethod]
    public void GetRows_ArePopulatedCorrectly() {
      events.Rows.Count.Should().Be(7);
      events.Rows[3].Cells[0].Value.Should().Be("Test Event 2");
      events.Rows[6][7].Should().Be("04 Time of Day", "because it implicitly converts to string");
      events.Rows[3][30].Should().Be("0");
    }

    [TestMethod]
    public void GetColumns_ArePopulatedCorrectly() {
      events.Columns.Count.Should().Be(31);
      events.Columns[7].Cells[2].Value.Should().Be("04 Time of Day");
      events.Columns[8][6].Should().Be("1:45 PM", "because it implicitly converts to string");
    }

    [TestMethod]
    public void SetFillerCellInColumn_CellIsInserted() {
      var dataElements = events.Rows[2].RowXML.Descendants(Settings.ssPrefix + "Data").ToList();
      dataElements[3].Value.Should().Be("1", "because this is cell 4 as the XML skipped an empty cell using an index attribute");
      events.Columns[3][2].Should().Be("");
      events.Columns[3].Cells[2].IsFillerCell.Should().Be(true);

      events.Columns[3].Cells[2].Value = "New Value";

      events.Columns[3].Cells[2].IsFillerCell.Should().Be(false);
      dataElements = events.Rows[2].RowXML.Descendants(Settings.ssPrefix + "Data").ToList();
      dataElements[4].Value.Should().Be("1", "because a new cell has been inserted since the blank cell now has a value");
      dataElements[3].Value.Should().Be("New Value", "because we added this cell when a value was assigned to it");
    }
  }
}
