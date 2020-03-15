using System;
using System.Text;
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
        private List<DataTypes.Quote> _data;
        private readonly TradeStrategyLib.Models.IStrategy _strategy;
        private readonly CultureInfo _culture = new CultureInfo("fr-FR");

        public StrategyBacktesterTester()
        {
        }

        /// <summary>
        ///Obtient ou définit le contexte de test qui fournit
        ///des informations sur la série de tests active, ainsi que ses fonctionnalités.
        ///</summary>
        public TestContext TestContext { get; set; }

        [ClassInitialize]
        public void InitClass(TestContext testContext)
        {
            // load market data for the tests
            var startDate = DateTime.Parse("02/02/2019", this._culture);
            var endDate = DateTime.Parse("02/02/2020", this._culture);
            DataImporter.DataImporter.Instance.ImportData("MSFT", startDate, endDate);
            this._data = DataImporter.DataImporter.Instance.GetData();
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

        [TestMethod]
        public void TestBacktest(TradeStrategyLib.Models.IStrategy strategy)
        {
            //
            // TODO: ajoutez ici la logique du test
            //
        }
    }
}
