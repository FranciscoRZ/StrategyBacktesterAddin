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
        private string _ticker; 
        private DateTime _startDate;
        private DateTime _endDate;
        private CultureInfo _culture = new CultureInfo("fr-FR");

        public void OnImportDataPress(CustomUI.IRibbonControl control)
        {
            // Get instance of importer
            var importer = DataImporter.DataImporter.Instance;
            importer.ImportData(_ticker, _startDate, _endDate);
            //Model.TimeSeriesEntry[] data = importer.GetData();
            var data = importer.GetData();
            DataWriter.WriteStockData(_ticker, data);
        }

        public void GetTickerValue(CustomUI.IRibbonControl control, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                this._ticker = text;
            }
        }

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
    }
}