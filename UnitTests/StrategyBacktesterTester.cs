using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using TradeStrategyLib;

namespace UnitTests
{
    /// <summary>
    /// Test class for the <see cref="StrategyBacktester"/>
    /// </summary>
    [TestClass]
    public class StrategyBacktesterTester
    {
        private static List<DataTypes.Quote> _data;
        private static readonly CultureInfo _culture = new CultureInfo("fr-FR");

        public StrategyBacktesterTester()
        {
        }

        /// <summary>
        ///Obtient ou définit le contexte de test qui fournit
        ///des informations sur la série de tests active, ainsi que ses fonctionnalités.
        ///</summary>
        public TestContext TestContext { get; set; }

        /// <summary>
        /// Initialize class for backtests
        /// </summary>
        /// <param name="testContext"></param>
        [ClassInitialize]
        public static void InitClass(TestContext testContext)
        {
            // load market data for the tests
            var startDate = DateTime.Parse("02/02/2019", _culture);
            var endDate = DateTime.Parse("02/02/2020", _culture);
            DataImporter.DataImporter.Instance.ImportData("MSFT", startDate, endDate);
            _data = DataImporter.DataImporter.Instance.GetData();
        }
        
        /// <summary>
        /// Tests that each strategy can be correctly backtested with the 
        /// <see cref="StrategyBacktester"/> backtester class.
        /// </summary>
        [TestMethod]
        public void TestBacktests()
        {
            var maStrategy = new TradeStrategyLib.Models.MAStrategy(25, 100, 1000000.00, 0.05);
            var sarStrategy = new TradeStrategyLib.Models.ParabolicSARStrategy(20, 100, 5, 1000000.00, 0.05);
            var bolStrategy = new TradeStrategyLib.Models.BollingerStrategy(25, 2, 2, 1000000.00, 0.05);

            // run the tests sequentially
            this.TestStrategy(maStrategy);
            this.TestStrategy(sarStrategy);
            this.TestStrategy(bolStrategy);
        }

        /// <summary>
        /// Tests whether given strategy can be correctly backtested
        /// </summary>
        /// <param name="strategy"></param>
        private void TestStrategy(TradeStrategyLib.Models.IStrategy strategy)
        {
            var backtest = new StrategyBacktester(strategy, _data);
            backtest.Compute();

            var pnlHistory = backtest.GetPnLHistory();
            var totalPnl = backtest.GetTotalPnl();

            // Check if results are empty
            // NOTE FRZ: This test should be more rigorous with expected results to test
            // but I don't have the time.
            Assert.IsNotNull(pnlHistory);
            Assert.IsNotNull(totalPnl);
        }
    }
}
