using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net;

namespace DataImporter
{
    public sealed class DataImporter
    {
        private string _key;
        private List<DataTypes.Quote> _data;
        // Instance to used throughout application
        private static readonly DataImporter instance = new DataImporter();

        public List<DataTypes.Quote> GetData()
        {
            return this._data;
        }

        public void ImportData(string symbol, DateTime startDate, DateTime endDate)
        {
            ReadConfig();
            string apiKey = this._key;
            string request = $"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&apikey={apiKey}&outputsize=full&datatype=csv";
            try
            {
                string response = new WebClient().DownloadString(request);
                string[] dailyPrices = Regex.Split(response, "\n");

                var data = new List<DataTypes.Quote>();
                // First occurence is the headers, which we don't want
                var iterPrices = dailyPrices.Skip(1).ToArray();
                // Last occurence is empty string, which we don't want
                iterPrices = iterPrices.Take(iterPrices.Count() - 1).ToArray();
                
                foreach (string line in iterPrices)
                {
                    data.Add(DataTypes.QuoteDeserializer.DeserializeAlphaVantage(line));
                }

                IEnumerable<DataTypes.Quote> query = from price in data
                                                     where (DateTime.Compare(startDate.Date, price.TimeStamp.Date) <= 0 &&
                                                            DateTime.Compare(price.TimeStamp.Date, endDate.Date) <= 0)
                                                     select price;
                // Reverse the order to chronological
                query = query.Reverse();

                this._data = query.ToList();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void ReadConfig()
        {
            this._key = ConfigurationManager.AppSettings.Get("AVKey2");
        }

        // Implement Singleton pattern
        private DataImporter()
        {
        }

        static DataImporter()
        {
        }

        public static DataImporter Instance
        {
            get
            {
                return instance;
            }
        }
    }

}
