using System;
using System.Text;

using DataTypes;

namespace TradeStrategyLib.Models
{
    public class TradeSituation : ITradeSituation
    {
        #region Fields
        /// <summary>
        /// Unique trade situation ID used by the field consequently named.
        /// </summary>
        private static int _tradeSituationId;

        /// <summary>
        /// Field used to save the property value
        /// </summary>
        private readonly bool _isLongTrade;

        /// <summary>
        /// Init in constructor ONCE. Use static tadeSituationId field!
        /// </summary>
        private readonly int _id;

        /// <summary>
        /// Take profit set for this trade situation. When hit, the position closes.
        /// </summary>
        private readonly double _tpInBps;

        /// <summary>
        /// Quote reference for the opening price and further information
        /// </summary>
        private Quote _executedOpenQuote;

        /// <summary>
        /// Quote reference when the position is closed
        /// </summary>
        private Quote _executedCloseQuote;

        /// <summary>
        /// Field used to save the current position Opened/Closed status
        /// </summary>
        private bool _isClosed = true;

        /// <summary>
        /// Max draw down in basis points
        /// </summary>
        private double _maxDDinBps;

        /// <summary>
        /// PnL calculation in basis points
        /// </summary>
        private double pnlInBps;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes new instance of the <see cref="TradeSituation"/> class.
        /// </summary>
        /// <param name="openOrder">Reference opening order</param>
        /// <param name="isLongTrade">True if the position is Long</param>
        /// <param name="tpInBps">Take profit set in basis points</param>
        public TradeSituation(Quote openOrder, bool isLongTrade, double tpInBps)
        {
            this._id = StaticTradeSituationId;
            this._isLongTrade = isLongTrade;
            if (tpInBps < 0.00)
            {
                throw new ArgumentException("The take profit must be positive");
            }
            this._tpInBps = tpInBps;
            this.OpenPosition(openOrder);
        }
        #endregion

        #region ImplementInterface
        /// <summary>
        /// Gets the unique trade situation ID
        /// </summary>
        public int GetTradeSituationId
        {
            get { return this._id; }
        }

        /// <summary>
        /// Gets the Opening quote
        /// </summary>
        public Quote GetExecutedOpenQuote
        {
            get { return this._executedOpenQuote; }
        }

        /// <summary>
        /// Gets the Closing quote
        /// </summary>
        public Quote GetExecutedCloseQuote
        {
            get { return this._executedCloseQuote; }
        }

        /// <summary>
        /// Gets the PnL of this position in basis points
        /// </summary>
        public double GetOrderPnlInBps
        {
            get
            {
                if (this._isClosed)
                {
                    if (this._executedCloseQuote == null)
                    {
                        throw new InvalidOperationException("Please initialize the closing orders list first");
                    }
                    return this.CalculatePnL();
                }
                else
                {
                    return this._maxDDinBps;
                }
            }
        }

        /// <summary>
        /// Gets the Price at which the position was opened
        /// </summary>
        public double GetEntryPrice
        {
            get { return this._executedOpenQuote.ClosePrice; }
        }

        /// <summary>
        /// Gets boolean indicator of whether the trade status:
        /// true => BUY
        /// false => SELL
        /// </summary>
        public bool IsLongTrade
        {
            get { return this._isLongTrade; }
        }

        /// <summary>
        /// Gets boolean indicator of whether position is closed
        /// </summary>
        public bool IsClosed
        {
            get { return this._isClosed; }
        }

        /// <summary>
        /// Gets the Maximum Draw Down statistic
        /// </summary>
        public double GetMaxDrawDown
        {
            get { return this._maxDDinBps; }
        }

        /// <summary>
        /// Procedure followed when the position is closed
        /// </summary>
        /// <param name="closingQuote">Reference quote to calculate all the statistics of the position</param>
        public void ClosePosition(Quote closingQuote)
        {
            this._isClosed = true;
            this._executedCloseQuote = closingQuote;
            this.CalculatePnL(closingQuote);
        }

        /// <summary>
        /// Procedure called when you create the object and open the position
        /// </summary>
        /// <param name="openingQuote">Reference quote saved as the 'open position' quote</param>
        public void OpenPosition(Quote openingQuote)
        {
            this._executedOpenQuote = openingQuote;
            this._isClosed = false;
        }

        /// <summary>
        /// Updates the pnl, maxdd and other values and parameters of this object.
        /// </summary>
        /// <param name="quote"></param>
        /// <returns><see cref="bool"/> True if the position if closed as a result of this new information.
        /// False otherwise</returns>
        public bool UpdateOnOrder(Quote quote)
        {
            this.CalculateMaxDD(quote);
            this.CalculatePnL(quote);
            if (this._isLongTrade && this._executedOpenQuote.ClosePrice * (1 + this._tpInBps / 10000) <= quote.ClosePrice)
            {
                // we bought and the position can be closed with this order
                this.ClosePosition(quote);
                return true;
            }
            else if (!this._isLongTrade && this._executedOpenQuote.ClosePrice * (1 - this._tpInBps / 10000) > quote.ClosePrice)
            {
                // we sold and the position can be closed with this order
                this.ClosePosition(quote);
                return true;
            }
            return false;
        }

        #endregion

        #region Protected Class Methods
        /// <summary>
        /// Gets the unique ID of this trade situation
        /// </summary>
        protected static int StaticTradeSituationId
        {
            get { return _tradeSituationId++; }
        }

        /// <summary>
        /// Calculates and updates the Max Draw Down property
        /// </summary>
        /// <param name="quote">Reference quote for Max DD calculation</param>
        /// <returns><see cref="double"/>Calculated max draw down if applicable</returns>
        protected double CalculateMaxDD(Quote quote)
        {
            // already closed? return the dd that was already calculated.
            if (this._isClosed)
            {
                return this._maxDDinBps;
            }

            if (this._isLongTrade)
            {
                // this is the order that can close the position (long position)
                if (quote.ClosePrice - this._executedOpenQuote.ClosePrice < this._maxDDinBps)
                {
                    // the situation is worse
                    this._maxDDinBps = quote.ClosePrice - this._executedOpenQuote.ClosePrice;
                }
            }
            else
            {
                // this is the order that can close the position (short position)
                if (this._executedOpenQuote.ClosePrice - quote.ClosePrice < this._maxDDinBps)
                {
                    // the situation is worse
                    this._maxDDinBps = this._executedOpenQuote.ClosePrice - quote.ClosePrice;
                }
            }
            return 0.00;
        }

        /// <summary>
        /// Calculates and updates the Profit and Loss variable
        /// </summary>
        /// <returns><see cref="double"/>Calculated PnL for a closed position</returns>
        protected double CalculatePnL()
        {
            if (this._isClosed)
            {
                this.pnlInBps = this.IsLongTrade ? (this._executedCloseQuote.ClosePrice - this._executedOpenQuote.ClosePrice) / this._executedOpenQuote.ClosePrice
                                                 : (this._executedOpenQuote.ClosePrice - this._executedCloseQuote.ClosePrice) / this._executedCloseQuote.ClosePrice;
            }
            else
            {
                throw new InvalidOperationException("Please initialize the closing orders list first.");
            }

            return this.pnlInBps;
        }

        /// <summary>
        /// Calculates and updates the Profit and Loss variable.
        /// </summary>
        /// <param name="quote">In regard of this quote</param>
        /// <returns><see cref="double"/> Calculated pnl</returns>
        protected double CalculatePnL(Quote quote)
        {
            this.pnlInBps = this._isLongTrade ? quote.ClosePrice - this._executedOpenQuote.ClosePrice
                                              : this._executedOpenQuote.ClosePrice - quote.ClosePrice;
            return this.pnlInBps;
        }

        #endregion
    }
}
