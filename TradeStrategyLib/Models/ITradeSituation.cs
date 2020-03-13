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
        int TradeSituationId { get; }

        /// <summary>
        /// Gets a list of orders executed when we bought / sold
        /// </summary>
        Quote ExecutedOpenQuote { get; }

    }
}
