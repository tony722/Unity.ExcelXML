using AET.Unity.ExcelXml;
using Crestron.SimplSharp.CrestronXmlLinq;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AET.Unity.ExcelXML.Tests {
  [TestClass]
  public class CellTests {
    Row row;
    //XElement rowXml;

    [TestInitialize]
    public void TestInit() {
      var cellsXml = XElement.Parse(SampleData.RowXml2);
      Settings.ssPrefix = cellsXml.GetNamespaceOfPrefix("ss");
      row = new Row(cellsXml, 31, null);
    }

    [TestMethod]
    public void Instantite_HasCorrectValue() {
      var cell = row.Cells[0];
      cell.Value.Should().Be("Test Event 3");
      string cellValue = cell;
      cellValue.Should().Be("Test Event 3", "because it converts to a string implicitly");
    }

    [TestMethod]
    public void NotIndexedCell_HasIsIndexedPropertyNot() {
      var cell = row.Cells[0];
      cell.IsIndexed.Should().Be(false);
    }

    [TestMethod]
    public void IndexedCell_HasIndexedPropertySet() {
      var cell = row.Cells[4];
      cell.IsIndexed.Should().Be(true);
    }

    [TestMethod]
    public void IndexedCell_SetIndex_CanChangeIndexValue() {
      var cell = row.Cells[9];
      cell.Index = 5;
      cell.Index.Should().Be(5);
    }
    [TestMethod]
    public void NonIndexedCell_SetIndex_AddsIndexedProperty() {
      var cell = row.Cells[3];
      cell.IsIndexed.Should().Be(false);
      cell.Index = 3;
      cell.Index.Should().Be(3);
      cell.IsIndexed.Should().Be(true);
    }


    [TestMethod]
    public void FillerCell_HasFillerPropertySet() {
      var cell = new Cell(row) { IsFillerCell = true };
      cell.IsFillerCell.Should().Be(true);
    }

    [TestMethod]
    public void FillerCell_SetIsFillerToFalse_FillerPropertyCleared() {
      var cell = new Cell(row) { IsFillerCell = true };
      cell.IsFillerCell = false;
      cell.IsFillerCell.Should().Be(false);
    }

    [TestMethod]
    public void NonFillerCell_SetIsFillerToTrue_FillerPropertySet() {
      var cell = row.Cells[0];
      cell.IsFillerCell = true;
      cell.IsFillerCell.Should().Be(true);
    }

    [TestMethod]
    public void IndexedCell_SetIndexedToFalse_IndexedPropertyRemoved() {
      var cell = row.Cells[9];
      cell.IsIndexed = false;
      cell.CellXml.HasAttribute(Settings.ssPrefix + "Index").Should().BeFalse();
    }


  }
}
