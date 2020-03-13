using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Interface used to create Trading Strategy calculators
    /// </summary>
    public interface IStrategy
    {
        /// <summary>
        /// Get the amount of the strategy
        /// </summary>
        double StrategyAmount { get; }

        List<ITradeSituation> TradeSituationHistory { get; }
    }
}
