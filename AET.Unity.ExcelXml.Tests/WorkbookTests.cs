using AET.Unity.ExcelXml;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AET.Unity.ExcelXML.Tests {
  [TestClass]

  public class WorkbookTests {
    static Workbook workbook;

    [TestInitialize]
    public void TestInit() {
      workbook = new Workbook();
      var xmlData = System.IO.File.ReadAllText("Schedule.xml");
      workbook.LoadXmlData(xmlData);
    }

    [TestMethod]
    public void LoadXmlData_WorkbookHasCorrectWorksheets() {
      workbook.Worksheets.Should().Contain(s => s.Name == "Settings");
      workbook.Worksheets["Events"].Should().NotBeNull();
      workbook["Settings"].Should().NotBeNull();
    }
  }
}
