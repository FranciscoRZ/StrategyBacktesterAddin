using System;
using System.Linq;
using System.Collections.Generic;

using DataTypes;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Base class for all strategies
    /// </summary>
    public abstract class StrategyBase : IStrategy
    {
        /// <summary>
        /// Take profit in basis points
        /// </summary>
        protected readonly double _tpInBps;

        /// <summary>
        /// Amount invested in strategy
        /// </summary>
        protected readonly double _amount;

        /// <summary>
        /// Entire history of executed trades
        /// </summary>
        protected readonly List<ITradeSituation> _tradeSituationHistory = new List<ITradeSituation>();

        /// <summary>
        /// Current way: True for Buy, False for Sell
        /// </summary>
        protected bool _currentWay;

        /// <summary>
        /// The current trade situation (used to store the open position)
        /// </summary>
        protected ITradeSituation _currentTradeSituation = null;


        /// <summary>
        /// Strategy's daily value since the first trade.
        /// </summary>
        protected List<double> _strategyDailyValue = new List<double>();

        /// <summary>
        /// Get the amount used to trade
        /// </summary>
        public double GetStrategyAmount
        {
            get { return this._amount; }
        }

        /// <summary>
        /// Gets all the trade situations that have happened
        /// </summary>
        public List<ITradeSituation> GetTradeSituationHistory
        {
            get { return this._tradeSituationHistory; }
        }
        
        /// <summary>
        /// Used to force close the position (example, EOD)
        /// </summary>
        /// <param name="quote"></param>
        public void ForceClosePosition(Quote quote)
        {
            if (this._currentTradeSituation != null)
                this._currentTradeSituation.ClosePosition(quote);
        }

        /// <summary>
        /// Computes and returns the total PnL of the strategy
        /// </summary>
        /// <returns><see cref="double"/>Total PnL of strategy</returns>
        public double GetPnL()
        {
            double pnlInBps = 0.00;
            foreach(ITradeSituation tradeSituation in this._tradeSituationHistory)
            {
                pnlInBps += tradeSituation.GetOrderPnlInBps;
            }
            return pnlInBps * this._amount;
        }

        /// <summary>
        /// Compute the strategy's maximum drawdown as defined by the single 
        /// greatest drawdown in the strategy's trades
        /// </summary>
        /// <returns><see cref="double"/>the strategy's Maximum Drawdown</returns>
        public double GetMaximumDrawdown()
        {
            return -this._tradeSituationHistory.Select(x => x.GetMaxDrawDown).Min();
        }
        
        /// <summary>
        /// Computes the standard deviation of an enumerable
        /// </summary>
        /// <returns><see cref="double"/>The standard deviation</returns>
        private double ComputeStandardDeviation(IEnumerable<double> data)
        {
            double std = 0.0;

            if (data.Any())
            {
                double mean = data.Average();
                double sum = 0.0;
                foreach (double dot in data)
                {
                    sum += Math.Pow(dot - mean, 2);
                }
                std = sum / (data.Count() - 1);
                std = Math.Sqrt(std);
            }

            return std;
        }

        /// <summary>
        /// Compute the strategy's annualized volatility as the standard deviation of the 
        /// strategy's order's pnl
        /// </summary>
        /// <returns><see cref="double"/>The strategy's volatility</returns>
        public double GetStrategyVol()
        {
            //List<double> data = this._tradeSituationHistory.Select(x => x.GetOrderPnlInBps).ToList<double>();
            return Math.Sqrt(252) * ComputeStandardDeviation(this._strategyDailyValue);
        }

        /// <summary>
        /// Steps through a quote arrival in the strategy, and computes and updates the daily value of the strategy
        /// since the first trade.
        /// </summary>
        /// <param name="arrivedQuote"></param>
        /// <returns></returns>
        public bool Step(Quote arrivedQuote)
        {
            bool ret = this._step(arrivedQuote);
            double value;
            ITradeSituation lastTrade;
            // Update the strategy value
            if (this._tradeSituationHistory.Count != 0) // check that a trade has actually happened
            {
                if (this._currentTradeSituation == null) // there's no open position
                {
                    lastTrade = this._tradeSituationHistory.Last();
                    this._strategyDailyValue.Add(lastTrade.GetOrderPnlInBps);
                }
                else // there's an open position
                {
                    value = this._currentTradeSituation.IsLongTrade ? (arrivedQuote.ClosePrice - this._currentTradeSituation.GetEntryPrice)
                                                                      / this._currentTradeSituation.GetEntryPrice 
                                                                    : (this._currentTradeSituation.GetEntryPrice - arrivedQuote.ClosePrice)
                                                                      / arrivedQuote.ClosePrice;
                    this._strategyDailyValue.Add(value);    
                }
            }
            return ret;
        }

        /// <summary>
        /// Computes the indicator and steps through the strategy (open or close the position)
        /// </summary>
        /// <param name="arrivedQuote">USed to update all the parameters for the trading strategy</param>
        /// <returns><see cref="bool"/> true if the position if opened (or flipped), false otherwise</returns>
        protected abstract bool _step(Quote arrivedQuote);

        /// <summary>
        /// Constructor for base asbtract class, called in child classes.
        /// </summary>
        /// <param name="tpInBps">Take Profit for the strategy in Bps.</param>
        /// <param name="amount">Amount invested in the strategy</param>
        protected StrategyBase(double tpInBps, double amount)
        {
            this._amount = amount;
            this._tpInBps = tpInBps;
        }
    }
}
