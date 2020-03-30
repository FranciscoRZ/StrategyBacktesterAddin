using System.Collections.Generic;

using DataTypes;

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
        double GetStrategyAmount { get; }
        
        /// <summary>
        /// Gets all the Trade Situations created so far. List of Trade Situations
        /// that exist so far.
        /// </summary>
        List<ITradeSituation> GetTradeSituationHistory { get; }

        /// <summary>
        /// Computes indicator and steps through strategy (open / close / nothing).
        /// </summary>
        /// <param name="arrivedQuote">Quote that has just arrived</param>
        /// <returns><see cref="bool"/> true: There is an action to take (buy or sell). False: no action to take  </returns>
        bool Step(Quote arrivedQuote);

        /// <summary>
        /// Force close any hanging position: in case there would be any open positions.
        /// Saves at the end of the day.
        /// </summary>
        /// <param name="quote">Quote used to perform calculations</param>
        void ForceClosePosition(Quote quote);

        /// <summary>
        /// Returns the total profit (or loss) of this investment strategy.
        /// </summary>
        /// <returns>Returns PnL of strategy</returns>
        double GetPnL();
    }
}
