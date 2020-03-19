using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Globalization;
using System.Runtime.InteropServices;

using CustomUI = ExcelDna.Integration.CustomUI;
using TradeStrategyLib;

namespace StrategyBacktesterAddin
{
    [ComVisible(true)]
    public class RibbonControler : CustomUI.ExcelRibbon
    {
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
            //Model.TimeSeriesEntry[] data = importer.GetData();
            var data = importer.GetData();
            DataWriter.WriteStockData(_ticker, data);
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
            List<DateTime> dates = backtest.GetDates();
            double totalPnl = backtest.GetTotalPnl();

            // Write the results
            DataWriter.WriteBacktestResults("Moving Average", totalPnl, pnlHistory, dates);
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

    }
}