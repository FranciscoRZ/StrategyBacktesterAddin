using DataImporter;

using System.Collections.Generic;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;


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
            object[] dates1D = queryDate.Cast<object>().Reverse().ToArray();
            object[,] dates = Make2DArray(dates1D);

            var queryOpen = from stock in data
                            select stock.OpenPrice;
            object[] open1D = queryOpen.Cast<object>().Reverse().ToArray();
            object[,] open = Make2DArray(open1D);

            var queryHigh = from stock in data
                            select stock.HighPrice;
            object[] high1D = queryHigh.Cast<object>().Reverse().ToArray();
            object[,] high = Make2DArray(high1D);

            var queryLow = from stock in data
                           select stock.LowPrice;
            object[] low1D = queryLow.Cast<object>().Reverse().ToArray();
            object[,] low = Make2DArray(low1D);

            var queryClose = from stock in data
                             select stock.ClosePrice;
            object[] close1D = queryClose.Cast<object>().Reverse().ToArray();
            object[,] close = Make2DArray(close1D);

            var queryVol = from stock in data
                           select stock.Volume;
            object[] volume1D = queryVol.Cast<object>().Reverse().ToArray();
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
    }
}