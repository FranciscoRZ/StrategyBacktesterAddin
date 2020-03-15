using System;
using System.Globalization;

namespace DataTypes
{
    /// <summary>
    /// Class for deserializing incoming market data
    /// </summary>
    public static class QuoteDeserializer
    {
        public static Quote DeserializeAlphaVantage(string input)
        {
            var quote = new Quote();
            string[] data = input.Split(',');
            quote.TimeStamp = DateTime.Parse(data[0]);
            quote.OpenPrice = double.Parse(data[1], CultureInfo.InvariantCulture);
            quote.HighPrice = double.Parse(data[2], CultureInfo.InvariantCulture);
            quote.LowPrice = double.Parse(data[3], CultureInfo.InvariantCulture);
            quote.ClosePrice = double.Parse(data[4], CultureInfo.InvariantCulture);
            quote.Volume = double.Parse(data[5].Replace("r",""), CultureInfo.InvariantCulture);
            return quote;
        }

    }
}
