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

        /// <summary>
        /// Set up the class for import before each test
        /// </summary>
        [TestInitialize]
        public void SetUp()
        {
            this._ticker = "MSFT";
            this._culture = new CultureInfo("fr-FR");
            this._startDate = DateTime.Parse("02/02/2019", this._culture);
            this._endDate = DateTime.Parse("02/02/2020", this._culture);
        }

        /// <summary>
        /// Tests wheter the data is correctly imported
        /// </summary>
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

        /// <summary>
        /// Tests whether the singleton pattern is well implemented.
        /// </summary>
        [TestMethod]
        public void TestSingletonPattern()
        {
            // get first reference to instance of Importer
            var firstRef = DataImporter.DataImporter.Instance;

            // get second reference to instance of Importer
            var secondRef = DataImporter.DataImporter.Instance;

            Assert.ReferenceEquals(firstRef.GetData(), secondRef.GetData());
        }
    }
}
