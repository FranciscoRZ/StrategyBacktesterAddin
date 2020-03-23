using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

using DataImporter;

namespace StrategyBacktesterAddin
{
    public static class DataWriter
    {
        private static object[,] Make2DArray(object[] input)
        {
            object[,] output = new object[input.GetLength(0), 1];
            for (int i = 0; i < input.GetLength(0); i++)
            {
                output[i, 0] = input[i];
            }
            return output;
        }

        public static void WriteStockData(string ticker, List<DataTypes.Quote> data)
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

        public static void WriteBacktestResults(string strategyName, double totalPnl,
                                                List<double> pnlHistory, List<DateTime> dates)
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

            // Show results
            XLSingleton.Instance.XLApp.ScreenUpdating = true;

        }
    }
}