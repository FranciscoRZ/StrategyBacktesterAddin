﻿using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Globalization;
using System.Runtime.InteropServices;

using CustomUI = ExcelDna.Integration.CustomUI;
using TradeStrategyLib;

namespace StrategyBacktesterAddin
{
    /// <summary>
    /// Class contains all the callback functions of the buttons in our ribbon
    /// </summary>   
    [ComVisible(true)]
    public class RibbonControler : CustomUI.ExcelRibbon
    {
        #region Class Properties
        // Import parameters
        private string _ticker;
        private DateTime _startDate;
        private DateTime _endDate;
        private CultureInfo _culture = new CultureInfo("fr-FR");

        // MA parameters
        private int? _maShortLevel = null;
        private int? _maLongLevel = null;
        private double? _maAmount = null;
        private double? _maTakeProfitInBps = null;

        // Bollinger parameters
        private int? _bolShortLevel = null;
        private double? _bolUpperBound = null;
        private double? _bolLowerBound = null;
        private double? _bolAmount = null;
        private double? _bolTakeProfitInBps = null;

        // Parabolic SAR parameters
        private double? _SARAccFactorLevel = null;
        private double? _SARMaxAccFactorLevel = null;
        private double? _SARAccFactorStep = null;
        private double? _SARAmount = null;
        private double? _SARTakeProfitInBps = null;
        #endregion

        #region Import Data
        /// <summary>
        /// Callback method for Import Data button.
        /// Gets the requested data for backtesting and writes it to a new worksheet.
        /// </summary>
        /// <param name="control">Exposes method to ribbon</param>
        public void OnImportDataPress(CustomUI.IRibbonControl control)
        {
            // Get instance of importer
            var importer = DataImporter.DataImporter.Instance;
            importer.ImportData(_ticker, _startDate, _endDate);
            
            // Get the data and write it
            var data = importer.GetData();
            DataWriter.WriteStockData(_ticker, data);

            // Ensure all COM objects are released
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Callback method for Ticker edit box.
        /// Inputs the value of the ticker for which data is imported.
        /// </summary>
        /// <param name="control">Exposes method to ribbon</param>
        /// <param name="text">Value in the edit box</param>
        public void GetTickerValue(CustomUI.IRibbonControl control, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                this._ticker = text;
            }
        }

        /// <summary>
        /// Callback method for Start Date edit box.
        /// Inputs the value in the box and casts it to DateTime object.
        /// </summary>
        /// <param name="control">Exposes method to ribbon</param>
        /// <param name="text">Value in the edit box</param>
        public void GetStartDateValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._startDate = DateTime.Parse(text, this._culture);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for End Date edit box.
        /// Inputs the value in the box and casts it to DateTime object.
        /// </summary>
        /// <param name="control">Exposes method to ribbon</param>
        /// <param name="text">Value in the edit box</param>
        public void GetEndDateValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._endDate = DateTime.Parse(text, this._culture);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }
        #endregion

        #region Moving Average

        /// <summary>
        /// Callback method for Launch MA Test button.
        /// Launches the MA strategy backtest.
        /// </summary>
        /// <param name="control">Exposes method to ribbon</param>
        public void TestMovingAverage(CustomUI.IRibbonControl control)
        {
            #region TestMovingAverage sanity checks
            // --> On imported data
            var importer = DataImporter.DataImporter.Instance;
            List<DataTypes.Quote> data = importer.GetData();
            if (data == null)
            {
                MessageBox.Show("Please import data before lauching test");
                return;
            }
            // --> On strategy parameters
            if (this._maAmount == null)
            {
                MessageBox.Show("Please input amount to invest in strategy");
                return;
            }
            if (this._maLongLevel == null)
            {
                MessageBox.Show("Please input Moving Average Long Level in strategy");
                return;
            }
            if (this._maShortLevel == null)
            {
                MessageBox.Show("Please input Moving Average Short Level parameter");
                return;
            }
            if (this._maTakeProfitInBps == null)
            {
                MessageBox.Show("Please input Moving Average Take Profit parameter");
                return;
            }
            #endregion

            // Compute the backtest and get the results
            var strategy = new TradeStrategyLib.Models.MAStrategy((int)this._maShortLevel, (int)this._maLongLevel,
                                                                  (double)this._maAmount, (double)this._maTakeProfitInBps);
            var backtest = new StrategyBacktester(strategy, data);
            backtest.Compute();

            List<double> pnlHistory = backtest.GetPnLHistory();
            if (!pnlHistory.Any())
            {
                MessageBox.Show("Moving average strategy did not generate any trades on this" +
                                " time interval");
                return;
            }

            List<DateTime> dates = backtest.GetDates();
            double totalPnl = backtest.GetTotalPnl();
            double? maxDD = backtest.GetMaximumDrawdown();
            double vol = backtest.GetStrategyVol();

            // Write the results
            DataWriter.WriteBacktestResults("Moving Average", totalPnl, maxDD, vol, pnlHistory, dates);

            // Ensure all COM objects are released
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Callback method for MA Short Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetMAShortLevelValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._maShortLevel = int.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for MA Long Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetMALongLevelValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._maLongLevel = int.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for MA Amount edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetMAAmountValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._maAmount = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for MA Take Profit Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetMATakeProfitValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._maTakeProfitInBps = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        #endregion

        #region Bollinger

        /// <summary>
        /// Callback method for Launch MA Test button.
        /// Launches the MA strategy backtest.
        /// </summary>
        /// <param name="control">Exposes method to ribbon</param>
        public void TestBollinger(CustomUI.IRibbonControl control)
        {
            #region TestBollinger sanity checks
            // --> On imported data
            var importer = DataImporter.DataImporter.Instance;
            List<DataTypes.Quote> data = importer.GetData();
            if (data == null)
            {
                MessageBox.Show("Please import data before lauching test");
                return;
            }
            // --> On strategy parameters
            if (this._bolAmount == null)
            {
                MessageBox.Show("Please input amount to invest in strategy");
                return;
            }
            if (this._bolShortLevel == null)
            {
                MessageBox.Show("Please input Moving Average Short Level parameter");
                return;
            }
            if (this._bolUpperBound == null)
            {
                MessageBox.Show("Please input Upper Bound coefficient parameter");
                return;
            }
            if (this._bolLowerBound == null)
            {
                MessageBox.Show("Please input Lower Bound coefficient parameter");
                return;
            }
            if (this._bolTakeProfitInBps == null)
            {
                MessageBox.Show("Please input Moving Average Take Profit parameter");
                return;
            }
            #endregion

            // Compute the backtest and get the results
            var bol_Strategy = new TradeStrategyLib.Models.BollingerStrategy((int)this._bolShortLevel, (double)this._bolUpperBound,
                                                                         (double)this._bolLowerBound, (double)this._bolAmount,
                                                                         (double)this._bolTakeProfitInBps);
            var bol_Backtest = new StrategyBacktester(bol_Strategy, data);
            bol_Backtest.Compute();

            List<double> bol_pnlHistory = bol_Backtest.GetPnLHistory();
            if (!bol_pnlHistory.Any())
            {
                MessageBox.Show("Bollinger strategy did not generate any trades on this " +
                                "time interval");
                return;
            }
            List<DateTime> bol_dates = bol_Backtest.GetDates();
            double bol_totalPnl = bol_Backtest.GetTotalPnl();
            double? maxDD = bol_Backtest.GetMaximumDrawdown();
            double vol = bol_Backtest.GetStrategyVol();

            // Write the results
            DataWriter.WriteBacktestResults("Bollinger", bol_totalPnl, maxDD, vol, bol_pnlHistory, bol_dates);

            // Ensure all COM objects are realeased
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Callback method for MA Short Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetBolShortLevelValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._bolShortLevel = int.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Bollinger Upper Bound edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetBolUpperBoundValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._bolUpperBound = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Bollinger Lower Bound edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetBolLowerBoundValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._bolLowerBound = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Bollinger Amount edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetBolAmountValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._bolAmount = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Bollinger Take Profit Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetBolTakeProfitValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._bolTakeProfitInBps = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        #endregion

        #region Parabolic SAR

        /// <summary>
        /// Callback method for Launch MA Test button.
        /// Launches the MA strategy backtest.
        /// </summary>
        /// <param name="control">Exposes method to ribbon</param>
        public void TestParabolicSAR(CustomUI.IRibbonControl control)
        {
            #region TestParabolicSAR sanity checks
            // --> On imported data
            var importer = DataImporter.DataImporter.Instance;
            List<DataTypes.Quote> data = importer.GetData();
            if (data == null)
            {
                MessageBox.Show("Please import data before lauching test");
                return;
            }
            // --> On strategy parameters
            if (this._SARAmount == null)
            {
                MessageBox.Show("Please input amount to invest in strategy");
                return;
            }
            if (this._SARAccFactorLevel == null)
            {
                MessageBox.Show("Please input Accelerator Factor parameter");
                return;
            }
            if (this._SARMaxAccFactorLevel == null)
            {
                MessageBox.Show("Please input Maximum Accelerator Factor parameter");
                return;
            }
            if (this._SARTakeProfitInBps == null)
            {
                MessageBox.Show("Please input SAR Take Profit parameter");
                return;
            }
            #endregion

            // Compute the backtest and get the results
            var Strategy = new TradeStrategyLib.Models.ParabolicSARStrategy((double)this._SARAccFactorLevel,
                                                                         (double)this._SARMaxAccFactorLevel,
                                                                         (double)this._SARAccFactorStep,
                                                                         (double)this._SARAmount,
                                                                         (double)this._SARTakeProfitInBps);
            var Backtest = new StrategyBacktester(Strategy, data);
            Backtest.Compute();

            List<double> pnlHistory = Backtest.GetPnLHistory();
            if (!pnlHistory.Any())
            {
                MessageBox.Show("Parabolic SAR strategy did not generate any trades " +
                                "in this time interval.");
                return;
            }
            List<DateTime> dates = Backtest.GetDates();
            double totalPnl = Backtest.GetTotalPnl();
            double? maxDD = Backtest.GetMaximumDrawdown();
            double vol = Backtest.GetStrategyVol();

            // Write the results
            DataWriter.WriteBacktestResults("Parabolic SAR", totalPnl, maxDD, vol, pnlHistory, dates);

            // Ensure COM objects are released
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Callback method for Accelerator Factor edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetSARAccFactorValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._SARAccFactorLevel = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Maximum Accelerator Factor edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetSARMaxAccFactorValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._SARMaxAccFactorLevel = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Accelerator Factor Step edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetSARAccFactorStepValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._SARAccFactorStep = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Parabolic SAR Amount edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetSARAmountValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._SARAmount = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Callback method for Parabolic SAR Take Profit Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        /// <exception cref="MessageBox">Shows message box on FormatException</exception>
        public void GetSARTakeProfitValue(CustomUI.IRibbonControl control, string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    this._SARTakeProfitInBps = double.Parse(text);
                }
            }
            catch (FormatException e)
            {
                MessageBox.Show(e.Message);
            }
        }

        #endregion

        /// <summary>
        /// Callback method for Launch All button.
        /// Creates Task objects for each strategy backtest and runs them
        /// asynchronously on a ThreadPool by calling the Run method.
        /// Method is safe as long as <see cref="XLSingleton"/> 
        /// and <see cref="DataWriter"/> implemenations remain thread-safe.
        /// </summary>
        /// <param name="control">Exposes the method to the ribbon</param>
        public void OnLaunchAllPress(CustomUI.IRibbonControl control)
        {
            // Launch all tasks
            Task movingAverageWorker = Task.Run(() => TestMovingAverage(control));
            Task parabolicSARWorker = Task.Run(() => TestParabolicSAR(control));
            Task bollingerWorker = Task.Run(() => TestBollinger(control));

            // Make sure tasks have been completed. If not, .GetResult() propagates
            // the original exception instead of the AggregatedExceptionResult usually
            // propragated by tasks
            var movingAverageAwaiter = movingAverageWorker.GetAwaiter();
            movingAverageAwaiter.OnCompleted(() =>
            {
                try
                {
                    movingAverageAwaiter.GetResult();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Moving Average calculation failed. Error message : " + e.Message);
                }

            });
            var parabolicSARAwaiter = parabolicSARWorker.GetAwaiter();
            parabolicSARAwaiter.OnCompleted(() =>
            {
                try
                {
                    parabolicSARAwaiter.GetResult();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Parabolic SAR calculation failed. Error message : " + e.Message);
                }
            });
            var bollingerAwaiter = bollingerWorker.GetAwaiter();
            bollingerAwaiter.OnCompleted(() =>
            {
                try
                {
                    bollingerAwaiter.GetResult();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Bollinger calculation failed. Error message : " + e.Message);
                    ;
                }
            });

            // Ensure all COM objects are released
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}