 using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace StrategyBacktesterAddin
{
    /// <summary>
    /// Static class for writing results to workbook
    /// </summary>
    public static class DataWriter
    {
        /// <summary>
        /// Adjusts input to shape Excel-Dna understands and can print
        /// </summary>
        /// <param name="input">1-D Input result for reshaping</param>
        /// <returns>2-D version of input</returns>
        private static object[,] Make2DArray(object[] input)
        {
            object[,] output = new object[input.GetLength(0), 1];
            for (int i = 0; i < input.GetLength(0); i++)
            {
                output[i, 0] = input[i];
            }
            return output;
        }

        /// <summary>
        /// Wrapper around _writeStockData to ensure release of COM objects
        /// Writes the OHLCV daily historical data that was imported to workbook
        /// in new worksheet
        /// </summary>
        /// <param name="ticker">Ticker for which data was imported</param>
        /// <param name="data">Data that was imported</param>
        public static void WriteStockData(string ticker, List<DataTypes.Quote> data)
        {
            _writeStockData(ticker, data);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Writes the OHLCV daily historical data that was imported to workbook
        /// in new worksheet
        /// </summary>
        /// <param name="ticker">Ticker for which data was imported</param>
        /// <param name="data">Data that was imported</param>
        private static void _writeStockData(string ticker, List<DataTypes.Quote> data)
        {
            // Set up workbook
            XLSingleton.Instance.XLApp.ScreenUpdating = false;

            // Set up worksheet
            Excel.Worksheet ws = XLSingleton.Instance.XLWorkbook.Worksheets.Add(Type: Excel.XlSheetType.xlWorksheet);
            ws.Name = ticker + " data";
            ws.Range["A1", "F1"].Font.Bold = true;
            ws.Range["A1", "F1"].Value2 = new string[] { "Date", "Open", "High", "Low", "Close", "Volume" };

            // Extract arrays from dataset and make 2D for printing
            var queryDate = from stock in data
                            select stock.TimeStamp.Date.Date;
            object[] dates1D = queryDate.Cast<object>().ToArray();
            object[,] dates = Make2DArray(dates1D);
            var queryOpen = from stock in data
                            select stock.OpenPrice;
            object[] open1D = queryOpen.Cast<object>().ToArray();
            object[,] open = Make2DArray(open1D);
            var queryHigh = from stock in data
                            select stock.HighPrice;
            object[] high1D = queryHigh.Cast<object>().ToArray();
            object[,] high = Make2DArray(high1D);
            var queryLow = from stock in data
                           select stock.LowPrice;
            object[] low1D = queryLow.Cast<object>().ToArray();
            object[,] low = Make2DArray(low1D);
            var queryClose = from stock in data
                             select stock.ClosePrice;
            object[] close1D = queryClose.Cast<object>().ToArray();
            object[,] close = Make2DArray(close1D);
            var queryVol = from stock in data
                           select stock.Volume;
            object[] volume1D = queryVol.Cast<object>().ToArray();
            object[,] volume = Make2DArray(volume1D);

            // Write to worksheet
            ws.Range["A2"].Resize[dates.GetLength(0), 1].Value2 = dates;
            ws.Range["B2"].Resize[open.GetLength(0), 1].Value2 = open;
            ws.Range["C2"].Resize[high.GetLength(0), 1].Value2 = high;
            ws.Range["D2"].Resize[low.GetLength(0), 1].Value2 = low;
            ws.Range["E2"].Resize[close.GetLength(0), 1].Value2 = close;
            ws.Range["F2"].Resize[volume.GetLength(0), 1].Value2 = volume;

            // Format columns
            ws.Range["A2"].Resize[dates.GetLength(0), 1].NumberFormat = "dd/mm/yyyy";
            ws.Range["B2"].Resize[ws.Range["A2"].End[Excel.XlDirection.xlDown].Row,
            ws.Range["A2"].End[Excel.XlDirection.xlToRight].Column - 1]
                          .NumberFormat = "#,##0.00";
            ws.Range["F2"].Resize[ws.Range["F2"].End[Excel.XlDirection.xlDown].Row, 1].NumberFormat = "#,##0";

            // Show results
            XLSingleton.Instance.XLApp.ScreenUpdating = true;
        }

        /// <summary>
        /// Wrapper for _writeBacktestResults to ensure COM objects are released
        /// Writes the results of a Backtests for a given trading strategy to the 
        /// workbook.
        /// </summary>
        /// <param name="strategyName">The backtested strategy</param>
        /// <param name="totalPnl">The total PnL of this strategy on this backtest</param>
        /// <param name="pnlHistory">The pnl history of the strategy by trade on this backtest</param>
        /// <param name="dates">The dates at which the trades were made </param>
        public static void WriteBacktestResults(string strategyName, double totalPnl, double maxDD, 
                                                double vol, List<double> pnlHistory, List<DateTime> dates)
        {
            _writeBacktestResults(strategyName, totalPnl, maxDD, vol, pnlHistory, dates);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Writes the results of a Backtests for a given trading strategy to the 
        /// workbook.
        /// </summary>
        /// <param name="strategyName">The backtested strategy</param>
        /// <param name="totalPnl">The total PnL of this strategy on this backtest</param>
        /// <param name="pnlHistory">The pnl history of the strategy by trade on this backtest</param>
        /// <param name="dates">The dates at which the trades were made </param>
        private static void _writeBacktestResults(string strategyName, double totalPnl,
                                                  double maxDD, double vol, List<double> pnlHistory, List<DateTime> dates)
        {
            // Set up workbook
            XLSingleton.Instance.XLApp.ScreenUpdating = false;
            Excel.Worksheet ws = XLSingleton.Instance.XLApp.Worksheets.Add(Type: Excel.XlSheetType.xlWorksheet);
            ws.Name = strategyName + " results";
            ws.Range["A1:D1"].Font.Bold = true;
            ws.Range["A1:D1"].Value2 = new string[] { $"{strategyName} backtest results", "Total PnL", "Date", "PnL History en %" };

            // Set up the results for writing
            object[] pnlHistory1D = pnlHistory.Cast<object>().ToArray();
            object[,] pnlHistory2D = Make2DArray(pnlHistory1D);
            object[] dates1D = dates.Cast<object>().ToArray();
            object[,] dates2D = Make2DArray(dates1D);

            // Write the results
            ws.Range["B2"].Value2 = (object)totalPnl;
            ws.Range["C2"].Resize[dates2D.GetLength(0), 1].Value2 = dates2D;
            ws.Range["D2"].Resize[pnlHistory2D.GetLength(0), 1].Value2 = pnlHistory2D;

            // Format the worksheet
            ws.Range["B2"].NumberFormat = "#,##0.00";
            ws.Range["C2"].Resize[dates2D.GetLength(0), 1].NumberFormat = "dd/mm/yyyy";
            ws.Range["D2"].Resize[pnlHistory2D.GetLength(0), 1].NumberFormat = "0.00%";

            // Write summary statistics and format
            ws.Range["B3"].Value2 = "Maximum Drawdown";
            ws.Range["B3"].Font.Bold = true;
            ws.Range["B4"].Value2 = (object)maxDD;
            ws.Range["B4"].NumberFormat = "#,##0.00";
            ws.Range["B5"].Value2 = "Volatility";
            ws.Range["B5"].Font.Bold = true;
            ws.Range["B6"].Value2 = (object)vol;
            ws.Range["B6"].NumberFormat = "0.00%";

            // Show results
            XLSingleton.Instance.XLApp.ScreenUpdating = true;

        }
    }
}