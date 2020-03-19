using System;

using System.Windows.Forms;
using System.Globalization;
using System.Runtime.InteropServices;
using CustomUI = ExcelDna.Integration.CustomUI;

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
        private double _maShortLevel;
        private double _maLongLevel;
        private double _maAmount;
        private double _maTakeProfitInBps;

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
                this._startDate = DateTime.Parse(text, this._culture);
            }
            catch (FormatException e)
            {
                if (!string.IsNullOrEmpty(text))
                {
                    MessageBox.Show(e.Message);
                }
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
                this._endDate = DateTime.Parse(text, this._culture);
            }
            catch (FormatException e)
            {
                if (!string.IsNullOrEmpty(text))
                {
                    MessageBox.Show(e.Message);
                }
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
            MessageBox.Show("Method not implemented");
        }

        /// <summary>
        /// Callback method for MA Short Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        public void GetMAShortLevelValue(CustomUI.IRibbonControl control, string text)
        {
            MessageBox.Show("Method not implemented");
        }

        /// <summary>
        /// Callback method for MA Long Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        public void GetMALongLevelValue(CustomUI.IRibbonControl control, string text)
        {
            MessageBox.Show("Method not implemented");
        }

        /// <summary>
        /// Callback method for MA Amount edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        public void GetMAAmountValue(CustomUI.IRibbonControl control, string text)
        {
            MessageBox.Show("Method not implemented");
        }

        /// <summary>
        /// Callback method for MA Take Profit Level edit box.
        /// Inputs the value in the edit box and casts it to double.
        /// </summary>
        /// <param name="control">Exposes method to the ribbon</param>
        /// <param name="text">Value in the edit box</param>
        public void GetMATakeProfitValue(CustomUI.IRibbonControl control, string text)
        {
            MessageBox.Show("Method not implemented");
        }

        #endregion

    }
}