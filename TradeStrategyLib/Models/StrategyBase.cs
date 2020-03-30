using System;
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
        /// Computes the indicator and steps through the strategy (open or close the position)
        /// </summary>
        /// <param name="arrivedQuote">USed to update all the parameters for the trading strategy</param>
        /// <returns><see cref="bool"/> true if the position if opened (or flipped), false otherwise</returns>
        public abstract bool Step(Quote arrivedQuote);

        protected StrategyBase(double tpInBps, double amount)
        {
            this._amount = amount;
            this._tpInBps = tpInBps;
        }
    }
}
