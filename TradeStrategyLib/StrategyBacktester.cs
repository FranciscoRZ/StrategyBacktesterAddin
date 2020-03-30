using System;
using System.Collections.Generic;
using System.Linq;

namespace TradeStrategyLib
{
    public class StrategyBacktester
    {
        /// <summary>
        /// The strategy to backtest
        /// </summary>
        private readonly Models.IStrategy _strategy;

        /// <summary>
        /// The market data to use
        /// </summary>
        private readonly List<DataTypes.Quote> _data;

        /// <summary>
        /// Constructor of the <see cref="StrategyBacktester"/> class.
        /// </summary>
        /// <param name="strategy">Strategy to backtest</param>
        /// <param name="data">Market data</param>
        public StrategyBacktester(Models.IStrategy strategy, List<DataTypes.Quote> data)
        {
            this._strategy = strategy;
            this._data = data;
        }

        /// <summary>
        /// Simulates the strategy for all the data given and computes statistics
        /// </summary>
        public void Compute()
        {
            for (int i = 0; i < this._data.Count() - 1; i++)
            {
                this._strategy.Step(this._data[i]);
            }
            this._strategy.ForceClosePosition(this._data.Last());
        }

        /// <summary>
        /// Gets all the dates in the trade history of the backtest
        /// </summary>
        /// <returns><see cref="List{DateTime}"/> of all dates in trade history</returns>
        public List<DateTime> GetDates()
        {
            var dates = new List<DateTime>();
            foreach(Models.ITradeSituation trade in this._strategy.GetTradeSituationHistory)
            {
                dates.Add(trade.GetExecutedCloseQuote.TimeStamp.Date);
            }
            return dates;
        }

        /// <summary>
        /// Gets the pnl history of the trade
        /// </summary>
        /// <returns><see cref="List{double}">List of the pnl of all the trades in backtest</returns>
        public List<double> GetPnLHistory()
        {
            List<Models.ITradeSituation> tradeHistory = this._strategy.GetTradeSituationHistory;
            var pnlHistory = new List<double>();
            foreach (Models.ITradeSituation trade in tradeHistory)
            {
                pnlHistory.Add(trade.GetOrderPnlInBps);
            }
            return pnlHistory;
        }

        /// <summary>
        /// Gets the total pnl of the strategy
        /// </summary>
        /// <returns><see cref="double"/> Total Pnl of the strategy</returns>
        public double GetTotalPnl()
        {
            return this._strategy.GetPnL();
        }
    }
}
