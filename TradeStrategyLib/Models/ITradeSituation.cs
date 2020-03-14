using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DataTypes;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Interface used to create Trade Situation calculators (perform Max Drawdown, PnL
    /// and other risk measures of position)
    /// </summary>
    public interface ITradeSituation
    {
        /// <summary>
        /// Gets a unique trade situation id, zero based. Increment it with every new Trade 
        /// Situation.
        /// </summary>
        int GetTradeSituationId { get; }

        /// <summary>
        /// Gets a list of orders executed when we bought / sold
        /// </summary>
        /// <returns>If we Buy returns offers (S). If we Sell returns Bids (B)</returns>
        Quote GetExecutedOpenQuote { get; }

        /// <summary>
        /// Gets the orders which were used to close the position
        /// </summary>
        /// <returns>List of NEW orders with the used orders to close the position.
        /// If we initially bought returns bids (B). If we initially sold retursn offers (S)</returns>
        Quote GetExecutedCloseQuote { get; }

        /// <summary>
        /// Gets the PnL.
        /// If the position is still open returns the worse PnL in basis points of the position.
        /// If the position is closed returns the closing value in BPS of the position.
        /// </summary>
        double GetOrderPnlInBps { get; }

        /// <summary>
        /// Gets the entry price.
        /// Double price. Holds the entry price. ex: 1.0002.
        /// </summary>
        double GetEntryPrice { get; }

        /// <summary>
        /// Gets a value indicating whether the trade is Long or Short
        /// Transaction way. True if we bought. False if we sold.
        /// </summary>
        bool IsLongTrade { get; }

        /// <summary>
        /// Gets a value indicating whether the position is closed.
        /// Check if the position was closed. True if the position was closed.
        /// False if the position is hanging.
        /// </summary>
        bool IsClosed { get; }

        /// <summary>
        /// Gets a value indicating whether the position is closed.
        /// </summary>
        double GetMaxDrawDown { get; }

        /// <summary>
        /// Upon fire of the UpdateOnOrder method -- use this method to save the orders that were used
        /// to close the position. All consequent calculations of the position should be done in this method.
        /// </summary>
        /// <param name="closingQuote">Order used to close the position</param>
        void ClosePosition(Quote closingQuote);

        /// <summary>
        /// You can chose to use this method to populate the position with orders when you create
        /// a new one. Or use the constructor to populate the position with orders.
        /// </summary>
        /// <param name="openingQuote">Order used to open the position</param>
        void OpenPosition(Quote openingQuote);

        /// <summary>
        /// Updates the statistics of the trade situation
        /// </summary>
        /// <param name="quote">Arrived order</param>
        /// <returns>True if the order vas closed. False if the order wasn't closed</returns>
        bool UpdateOnOrder(Quote quote);
    }
}
