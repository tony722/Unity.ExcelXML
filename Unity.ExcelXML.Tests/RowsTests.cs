using System.Linq;
using Crestron.SimplSharp.CrestronXmlLinq;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Unity.ExcelXml;

namespace Unity.ExcelXML.Tests {
  [TestClass]
  public class RowsTests {
    private Rows rows;
    private XElement rowsXml;

    [TestInitialize]
    public void TestInit() {
      rowsXml = XElement.Parse(SampleData.TableXml);
      Settings.ssPrefix = rowsXml.GetNamespaceOfPrefix("ss");
      rows = new Rows(rowsXml);
    }


    [TestMethod]
    public void RowsInstantiatedWithValidXML_HasCorrectNumberOfRows() {
      rows.Count.Should().Be(3);
    }

    [TestMethod]
    public void RowsInstantiatedWithValidXML_HasCorrectColumnsCount() {
      rows.ColumnsCount.Should().Be(31);
      rows[1].Cells.Count.Should().Be(31);
    }

    [TestMethod]
    public void ShortRow_HasCorrectColumnsCount() {
      rows[2].Cells.Count.Should().Be(31);
    }

    [TestMethod]
    public void Add_EmptyRowAdded() {
      rows.Add();
      rows.Count.Should().Be(4);
      rowsXml.Elements().Count().Should().Be(4);
    }

    [TestMethod]
    public void Add_CellsAreWritable() {
      rows.Add();
      rows.Count.Should().Be(4);
      rowsXml.Elements().Count().Should().Be(4);
      rows[3].Cells[3].Value = "test value";
      var newCellXml = rows[3].RowXML.Elements().First();
      newCellXml.Attribute(Settings.ssPrefix + "Index").Value.Should().Be("4", "because the first cell in this row is the one we just wrote to with an index of 4");
      newCellXml.Value.Should().Be("test value");      
    }

    [TestMethod]
    public void Remove_RowIsRemoved() {
      rows.Remove(rows.First());
      rows.Count.Should().Be(2, "because there were 3 sample rows before we removed this one.");
      rowsXml.Elements().Count().Should().Be(2);
    }
  }
}
