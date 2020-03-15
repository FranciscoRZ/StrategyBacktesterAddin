using System;
using System.Globalization;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
    [TestClass]
    public class DataImporterTester
    {
        private string _ticker;
        private DateTime _startDate;
        private DateTime _endDate;
        private CultureInfo _culture;

        [TestInitialize]
        public void SetUp()
        {
            this._ticker = "MSFT";
            this._culture = new CultureInfo("fr-FR");
            this._startDate = DateTime.Parse("02/02/2019", this._culture);
            this._endDate = DateTime.Parse("02/02/2020", this._culture);
        }

        [TestMethod]
        public void TestImportData()
        {
            // get instance of dataimporter
            var importerInstance = DataImporter.DataImporter.Instance;
            // load data
            importerInstance.ImportData(this._ticker, this._startDate, this._endDate);
            var data = importerInstance.GetData();

            Assert.IsNotNull(data);
        }
    }
}
