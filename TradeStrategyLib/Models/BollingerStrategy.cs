using System;
using System.Collections.Generic;
using System.Linq;
using TradeStrategyLib.Models;

using DataTypes;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Bollinger strategy 
    /// </summary>
    public class BollingerStrategy : IStrategy
    {
        /// <summary>
        /// Take profit in basis points
        /// </summary>
        private readonly double _tpInBps;
        
        /// <summary>
        /// Short level of moving average
        /// </summary>
        private readonly int _shortLevel;

        /// <summary>
        /// x / upper_bound = x*sigma:
        /// </summary>
        private readonly int _upperBound;

        /// <summary>
        /// x / lower_bound = -y*sigma:
        /// </summary>
        private readonly int _lowerBound;

        List<DataTypes.Quote> data;

        /// <summary>
        /// FIFO collection with a "short" history and all history
        /// </summary>
        private readonly FIFODoubleArray _shortPricesHistory;
        private readonly FIFODoubleArray _PricesHistory;

        /// <summary>
        /// The market data to use
        /// </summary>
        //private readonly List<DataTypes.Quote> _data;
        //private FIFODoubleArray _datatocalc;

        /// <summary>
        /// Trading amount
        /// </summary>
        private readonly double _amount;

        /// <summary>
        /// Private field of the TradeSituationHistory getter
        /// </summary>
        private readonly List<ITradeSituation> _tradeSituationHistory = new List<ITradeSituation>();

        /// <summary>
        /// Current Way: true for Buy, false for Sell
        /// </summary>
        private bool _currentWay;

        /// <summary>
        /// The current trade situation (used to store the open position mainly)
        /// </summary>
        private ITradeSituation _currentTradeSituation = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="MASrategy"/> class.
        /// Constructs a Moving Average strategy based on the long and short MA strategy
        /// parameters:
        /// If closing price < mean + std_inf * std => Buy
        /// If closing price > mean + std_sup * std => Sell
        /// </summary>
        /// <param name="maShortLevel">ex 25</param>
        /// <param name="maUpperBound">ex 2</param>
        /// <param name="maLowerBound">ex 2</param>
        /// <param name="amount">Amount invested in strategy</param>
        /// <param name="takeProfitInBps">PnL at which we take profit and close position</param>
        public BollingerStrategy(int bolShortLevel, int bolUpperBound, int bolLowerBound, double amount, double takeProfitInBps)
        {
            this._tpInBps = takeProfitInBps;
            this._shortLevel = bolShortLevel;
            this._upperBound = bolUpperBound;
            this._lowerBound = bolLowerBound;
            this._amount = amount;
            this._shortPricesHistory = new FIFODoubleArray(this._shortLevel);
            this._PricesHistory = new FIFODoubleArray(10000);
        }

        /// <summary>
        /// Get the amount used to trade
        /// </summary>
        /// <returns>The initial amount invested in strategy</returns>
        public double GetStrategyAmount
        {
            get { return this._amount;  }
        }

        /// <summary>
        /// Gets all the trade situations that have happened. 
        /// </summary>
        public List<ITradeSituation> GetTradeSituationHistory
        {
            get { return this._tradeSituationHistory; }
        }

        /// <summary>
        /// Computes the indicator and steps through the strategy (Open or Close position)
        /// </summary>
        /// <param name="arrivedQuote">Used to update all the parameters of the trading strategy.</param>
        /// <returns>true if the position is opened (or flipped), false otherwise</returns>
        public bool Step(Quote arrivedQuote)
        {
            // Update the data arrays
            this._shortPricesHistory.Put(arrivedQuote.ClosePrice);

            // Get the MAs for strategy
            double shortMean = this._shortPricesHistory.GetArrayMean();

            // std for all the short mobil average
            double std_MA = this._shortPricesHistory.GetArrayVol();

            // Histo prices
            this._PricesHistory.Put(arrivedQuote.ClosePrice);

            // std for all the period
            double std = this._PricesHistory.GetArrayVol();

            // mean for all the period
            double mean = this._PricesHistory.GetArrayMean();

            // Upper bound
            double up_bound = this._upperBound * std_MA;

            // Lower bound
            double lo_bound = this._lowerBound * std_MA;

            // Update the current position
            if (this._currentTradeSituation != null)
            {
                if (this._currentTradeSituation.UpdateOnOrder(arrivedQuote))
                {
                    // If true we closed the position
                    this._currentTradeSituation = null;
                }
            }

            // We flip if there's a change in position
            // Buy signal
            if (arrivedQuote.ClosePrice < mean + lo_bound * std && (!this._currentWay || // current way is sell
                                         this._currentTradeSituation == null)) // or there is no current trade situation
            {
                // close exiting order
                if (this._currentTradeSituation != null)
                {
                    this._currentTradeSituation.ClosePosition(arrivedQuote);
                }

                // Open position
                this._currentWay = true;
                this._currentTradeSituation = new TradeSituation(arrivedQuote, true, this._tpInBps);
                this._tradeSituationHistory.Add(this._currentTradeSituation);
                return true;
            }

            // Sell signal
            else if (arrivedQuote.ClosePrice > mean + up_bound * std && (this._currentWay || // current way is buy
                                                  this._currentTradeSituation == null)) // or there is no current trade situation
            {
                // Close existing order
                if (this._currentTradeSituation != null)
                {
                    this._currentTradeSituation.ClosePosition(arrivedQuote);
                }

                // Open position
                this._currentWay = false;
                this._currentTradeSituation = new TradeSituation(arrivedQuote, false, this._tpInBps);
                this._tradeSituationHistory.Add(this._currentTradeSituation);
                return true;
            }
            return false;
        }
        
        /// <summary>
        /// Used to force close the position (example, end-of-day)
        /// </summary>
        /// <param name="closingQuote">Reference Quote</param>
        public void ForceClosePosition(Quote closingQuote)
        {
            this._currentTradeSituation.ClosePosition(closingQuote);
        }

        /// <summary>
        /// Gets the Pnl of the strategy
        /// </summary>
        /// <returns></returns>
        public double GetPnL()
        {
            double pnlInBps = 0.00;
            foreach (ITradeSituation tradeSituation in this._tradeSituationHistory)
            {
                pnlInBps += tradeSituation.GetOrderPnlInBps;
            }

            return pnlInBps * this._amount;
        }

        /// <summary>
        /// Outputs all the calculations
        /// </summary>
        /// <returns></returns>
        public string OutputStrategyCalculations()
        {
            throw new NotImplementedException();
        }

    }
}
