using System;
using System.Globalization;

namespace DataImporter
{
    public class AlphaVantageData
    {
        public DateTime Timestamp { get; set; }
        public decimal Open { get; set; }

        public decimal High { get; set; }
        public decimal Low { get; set; }

        public decimal Close { get; set; }
        public decimal Volume { get; set; }

        public AlphaVantageData(string input)
        {
            string[] line = input.Split(',');
            this.Timestamp = DateTime.Parse(line[0]);
            this.Open = decimal.Parse(line[1], CultureInfo.InvariantCulture);
            this.High = decimal.Parse(line[2], CultureInfo.InvariantCulture);
            this.Low = decimal.Parse(line[3], CultureInfo.InvariantCulture);
            this.Close = decimal.Parse(line[4], CultureInfo.InvariantCulture);
            this.Volume = decimal.Parse(line[5].Replace("r", ""), CultureInfo.InvariantCulture);
        }
    }
}
