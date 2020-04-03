using System.Collections.Generic;
using DataTypes;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Moving Average (MA) strategy 
    /// </summary>
    public class MAStrategy : StrategyBase
    {
        /// <summary>
        /// Short level of moving average
        /// </summary>
        private readonly int _shortLevel;

        /// <summary>
        /// Long level of moving average
        /// </summary>
        private readonly int _longLevel;

        /// <summary>
        /// FIFO collection with a "long" history
        /// </summary>
        private readonly FIFODoubleArray _longPricesHistory;

        /// <summary>
        /// FIFO collection with a "short" history
        /// </summary>
        private readonly FIFODoubleArray _shortPricesHistory;

        /// <summary>
        /// List with the entire history of prices
        /// </summary>
        private List<double> _priceHistory = new List<double>();

        /// <summary>
        /// Initializes a new instance of the <see cref="MASrategy"/> class.
        /// Constructs a Moving Average strategy based on the long and short MA strategy
        /// parameters:
        /// If MA short > MA long => Buy
        /// If MA long > MA short => Sell
        /// </summary>
        /// <param name="maShortLevel">ex 25</param>
        /// <param name="maLongLevel">ex 100</param>
        /// <param name="amount">Amount invested in strategy</param>
        /// <param name="takeProfitInBps">PnL at which we take profit and close position</param>
        public MAStrategy(int maShortLevel, int maLongLevel, double amount, double takeProfitInBps)
               :base(takeProfitInBps, amount)
        {
            this._shortLevel = maShortLevel;
            this._longLevel = maLongLevel;
            this._longPricesHistory = new FIFODoubleArray(this._longLevel);
            this._shortPricesHistory = new FIFODoubleArray(this._shortLevel);
        }

        /// <summary>
        /// Computes the indicator and steps through the strategy (Open or Close position)
        /// </summary>
        /// <param name="arrivedQuote">Used to update all the parameters of the trading strategy.</param>
        /// <returns>true if the position is opened (or flipped), false otherwise</returns>
        protected override bool _step(Quote arrivedQuote)
        {
            // Update the data arrays
            this._longPricesHistory.Put(arrivedQuote.ClosePrice);
            this._shortPricesHistory.Put(arrivedQuote.ClosePrice);
            this._priceHistory.Add(arrivedQuote.ClosePrice);

            // Get the MAs for strategy
            double shortMean = this._shortPricesHistory.GetArrayMean();
            double longMean = this._longPricesHistory.GetArrayMean();

            // Update the current position
            if (this._currentTradeSituation != null)
            {
                if (this._currentTradeSituation.UpdateOnOrder(arrivedQuote))
                {
                    // If true we closed the position
                    this._currentTradeSituation = null;
                }
            }

            // Check if there's enough data for first signal
            if (this._priceHistory.Count >= this._longLevel)
            {
                // We flip if there's a change in position
                if (shortMean > longMean && (!this._currentWay || // current way is sell
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
                else if (shortMean < longMean && (this._currentWay ||
                                                  this._currentTradeSituation == null))
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
            }
            
            return false;
        }        
    }
}
