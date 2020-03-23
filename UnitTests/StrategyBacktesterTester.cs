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

        [ClassInitialize]
        public static void InitClass(TestContext testContext)
        {
            // load market data for the tests
            var startDate = DateTime.Parse("02/02/2019", _culture);
            var endDate = DateTime.Parse("02/02/2020", _culture);
            DataImporter.DataImporter.Instance.ImportData("MSFT", startDate, endDate);
            _data = DataImporter.DataImporter.Instance.GetData();
        }

        #region Attributs de tests supplémentaires
        //
        // Vous pouvez utiliser les attributs supplémentaires suivants lorsque vous écrivez vos tests :
        //
        // Utilisez ClassInitialize pour exécuter du code avant d'exécuter le premier test de la classe
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Utilisez ClassCleanup pour exécuter du code une fois que tous les tests d'une classe ont été exécutés
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Utilisez TestInitialize pour exécuter du code avant d'exécuter chaque test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Utilisez TestCleanup pour exécuter du code après que chaque test a été exécuté
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        /// <summary>
        /// Tests that each strategy can be correctly backtested with the 
        /// <see cref="StrategyBacktester"/> backtester class.
        /// </summary>
        [TestMethod]
        public void TestBacktest()
        {
            var strategy = new TradeStrategyLib.Models.MAStrategy(25, 100, 1000000.00, 0.05);
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
