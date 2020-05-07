using System;
using System.Collections.Generic;
using System.Linq;
using Crestron.SimplSharp.CrestronXmlLinq;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Unity.ExcelXml;

namespace Unity.ExcelXML.Tests {

  [TestClass]
  public class CellsTests {

    private XElement cellsXml;
    private Rows rows;
    private Row row;
    private Cells cells;

    [TestInitialize]
    public void TestInit() {
      cellsXml = XElement.Parse(SampleData.RowXml2);
      Settings.ssPrefix = cellsXml.GetNamespaceOfPrefix("ss");
      rows = new Rows();
      row = new Row(cellsXml, 31, rows);
      cells = row.Cells;
    }

    [TestMethod]
    public void TwoCells_CellsHaveCorrectValues() {   
      cells[0].Value.Should().Be("Test Event 3");
      cells[1].Value.Should().Be("3");
    }

    [TestMethod]
    public void CellWithIndex_CellsHaveCorrectValues() {
      cells[0].Value.Should().Be("Test Event 3");
      cells[3].Value.Should().Be("");
      cells[4].Value.Should().Be("1");
    }

    [TestMethod]
    public void StringIndexer_NoColumnNamesHaveBeenSet_ReturnsEmpty() {
      Action act = () => { var t = cells["oops"]; };
      act.Should().Throw<Exception>()
        .WithMessage("Column names have not been set. Cannot lookup column by name \"oops\".");
    }

    [TestMethod]
    public void StringIndexer_ValidColumnName_ReturnsValue() {
      rows.SetColumnNames(new List<string>(){"Col1","Col2"});
      cells["Col1"].Value.Should().Be("Test Event 3");
    }

    [TestMethod]
    public void StringIndexer_ValidColumnNameButWrongCase_ReturnsValue() {
      rows.SetColumnNames(new List<string>() { "Col1", "Col2" });
      cells["cOl1"].Value.Should().Be("Test Event 3");
    }

    [TestMethod]
    public void StringIndexer_InvalidColumnName_ReturnsEmptyString() {
      rows.SetColumnNames(new List<string>() { "Col1", "Col2" });
      cells["Oops"].Value.Should().Be(""); }

    [TestMethod]
    public void ModifyCellValue_XElementIsModified() {
      cells[0].Value = "New Value";
      var data = cellsXml.Descendants(Settings.ssPrefix + "Data").ToList();
      data[0].Value.Should().Be("New Value");
    }

    [TestMethod]
    public void ModifyAFillerCell_XElementIsAdded() {
      cells[10].Value = "New Value";
      var data = cellsXml.Descendants(Settings.ssPrefix + "Data").ToList();
      data[8].Value.Should().Be("New Value","because there are two skipped cells");
    }
  }
}
