using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DataTypes;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Moving Average (MA) strategy 
    /// </summary>
    public class MASrategy : IStrategy
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
        private ITradeSituation currentTradeSituation = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="MASrategy"/> class.
        /// Constructs a Moving Average strategy based on the long and short MA strategy
        /// parameters:
        /// If MA short > MA long => Buy
        /// If MA long > MA short => Sell
        /// </summary>
        /// <param name="maShortLevel"></param>
        /// <param name="maLongLevel"></param>
        /// <param name="amount"></param>
        /// <param name="takeProfitInBps"></param>
        public MASrategy(int maShortLevel, int maLongLevel, double amount, double takeProfitInBps)
        {
            this._tpInBps = takeProfitInBps;
            this._shortLevel = maShortLevel;
            this._longLevel = maLongLevel;
            this._amount = amount;
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
        /// Calculate if the position has to be opened
        /// </summary>
        /// <param name="arrivedQuote">Used to update all the parameters of the trading strategy.</param>
        /// <returns>true if the position is opened (or flipped), false otherwise</returns>
        public bool ShouldTakeAction(Quote arrivedQuote)
        {
            // Update the data arrays
            this._longPricesHistory.Put(arrivedQuote.ClosePrice);
            this._shortPricesHistory.Put(arrivedQuote.ClosePrice);

            // Get the MAs for strategy
            double shortMean = this._shortPricesHistory.GetArrayMean();
            double longMean = this._longPricesHistory.GetArrayMean();

            // We flip if there's a change in position
            if (shortMean > longMean && (!this._currentWay || // current way is sell
                                         this.currentTradeSituation == null)) // or there is no current trade situation
            {
                if (this.currentTradeSituation != null)
                {
                    this.currentTradeSituation.ClosePosition(arrivedQuote);
                }

                // Open position
                this._currentWay = true;
                this.currentTradeSituation = new TradeSituation(arrivedQuote, true, this._tpInBps);
                this._tradeSituationHistory.Add(this.currentTradeSituation);
                return true;
            }
            else if (shortMean < longMean && (this._currentWay|| 
                                              this.currentTradeSituation == null))
            {
                // Close existing order
                if (this.currentTradeSituation != null)
                {
                    this.currentTradeSituation.ClosePosition(arrivedQuote);
                }

                // Open position
                this._currentWay = false;
                this.currentTradeSituation = new TradeSituation(arrivedQuote, false, this._tpInBps);
                this._tradeSituationHistory.Add(this.currentTradeSituation);
                return true;
            }
            return false;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="closingQuote"></param>
        public void ForceClosePosition(Quote closingQuote)
        {
            this.currentTradeSituation.ClosePosition(closingQuote);
        }

        public double GetPnL()
        {
            throw new NotImplementedException();
        }

        public string OutputStrategyCalculations()
        {
            throw new NotImplementedException();
        }

    }
}
