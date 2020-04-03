using System;
using System.Linq;
using System.Collections.Generic;

using DataTypes;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Moving Average (MA) strategy 
    /// </summary>
    public class ParabolicSARStrategy : StrategyBase
    {
        /// <summary>
        /// Accelerator Factor a
        /// </summary>
        private readonly double _sarAccFactorLevel;

        /// <summary>
        /// Maximum Accelerator Factor
        /// </summary>
        private readonly double _sarMaxAccFactorLevel;

        /// <summary>
        /// Accelerator Factor Step
        /// </summary>
        private readonly double _sarAccFactorStep;

        /// <summary>
        /// FIFO collection with a SAR history with SAR(i-1) and SAR(i)
        /// </summary>
        //private readonly FIFODoubleArray _sarHistory;
        private List<double> _sarHistory = new List<double>();

        /// <summary>
        /// FIFO collection with a Price history with Price(i-1) and Price(i)
        /// </summary>
        //private readonly FIFODoubleArray _sarPricesHistory;
        private List<double> _sarPricesHistory = new List<double>();

        /// <summary>
        /// FIFO collection with a all data to calculate the Extremum Point EP
        /// </summary>
        private List<double> _epData = new List<double>();

        /// <summary>
        /// Extremum Point computed then stored in _epData
        /// </summary>
        private double _ep;

        /// <summary>
        /// Accelerator factor (a) to use from initial a to max a
        /// </summary>
        private List<double> _accFactor = new List<double>();

        /// <summary>
        /// SAR Formula
        /// </summary>
        private double _sar;

        /// <summary>
        /// Initializes a new instance of the <see cref="MASrategy"/> class.
        /// Constructs a Moving Average strategy based on the long and short MA strategy
        /// parameters:
        /// If SAR < Price => Buy
        /// If SAR > Price => Sell
        /// </summary>
        /// <param name="SARAccFactor">ex 0.02</param>
        /// <param name="SARMaxAccFactor">ex 0.2</param>
        /// /// <param name="SARAccFactorStep">ex 0.01</param>
        /// <param name="amount">Amount invested in strategy</param>
        /// <param name="takeProfitInBps">PnL at which we take profit and close position</param>
        public ParabolicSARStrategy(double SARAccFactor, double SARMaxAccFactor, double SARAccFactorStep,
                                    double SARamount, double takeProfitInBps)
               : base(takeProfitInBps, SARamount)
        {
            this._sarAccFactorLevel = SARAccFactor;
            this._sarMaxAccFactorLevel = SARMaxAccFactor;
            this._sarAccFactorStep = SARAccFactorStep;
        }

        /// <summary>
        /// Computes the indicator and steps through the strategy (Open or Close position)
        /// </summary>
        /// <param name="arrivedQuote">Used to update all the parameters of the trading strategy.</param>
        /// <returns>true if the position is opened (or flipped), false otherwise</returns>
        protected override bool _step(Quote arrivedQuote)
        {
            // Update the data arrays
            this._sarPricesHistory.Add(arrivedQuote.ClosePrice);

            // Fisrt SAR Formula 
            if (this._sarHistory.Count == 0) // Initialisation
            {
                this._sar = arrivedQuote.ClosePrice;
                this._sarHistory.Add(this._sar);
                this._accFactor.Add(this._sarAccFactorLevel);
            }

            // First computation for the Extremum Point
            else if (this._sarHistory.Count == 1)
            {
                this._ep = this._sarPricesHistory[this._sarPricesHistory.Count - 1];
                this._epData.Add(this._ep);

                // Get the SAR formula
                this._sar = this._sarPricesHistory[this._sarPricesHistory.Count - 2] +
                            this._sarAccFactorLevel * (this._ep - this._sarPricesHistory[this._sarPricesHistory.Count - 1]);
                this._sarHistory.Add(this._sar);
            }

            // Extremum Point calculation
            else
            {
                if (this._sarHistory[this._sarHistory.Count - 1] < this._sarPricesHistory[this._sarPricesHistory.Count - 2])
                {

                    // Same trend
                    if (this._sarHistory[this._sarHistory.Count - 2] <= this._sarPricesHistory[this._sarPricesHistory.Count - 3])
                    {
                        this._ep = Math.Max(this._sarPricesHistory[this._sarPricesHistory.Count - 1], this._epData.Max());
                        this._epData.Add(this._ep);
                    }

                    // Different trend
                    else
                    {
                        this._ep = this._sarPricesHistory[this._sarPricesHistory.Count - 1];
                        this._epData.Clear();
                        this._epData.Add(this._ep);
                        // Accelerator Factor incrementation
                        if (this._sarAccFactorLevel + this._sarAccFactorStep <= this._sarMaxAccFactorLevel)
                        {
                            this._accFactor.Add(this._sarAccFactorLevel + this._sarAccFactorStep);
                        }
                    }

                    // Get the SAR formula
                    this._sar = this._sarHistory[this._sarHistory.Count - 1] +
                                this._accFactor.Max() * (this._epData.Max() - this._sarHistory[this._sarHistory.Count - 1]);
                    this._sarHistory.Add(this._sar);
                }

                else
                {
                    // Same trend
                    if (this._sarHistory[this._sarHistory.Count - 2] >= this._sarPricesHistory[this._sarPricesHistory.Count - 3])
                    {
                        this._ep = Math.Min(this._sarPricesHistory[this._sarPricesHistory.Count - 1], this._epData.Min());
                        this._epData.Add(this._ep);
                    }
                    // Different trend
                    else
                    {
                        this._ep = this._sarPricesHistory[this._sarPricesHistory.Count - 1];
                        this._epData.Clear();
                        this._epData.Add(this._ep);
                        // Accelerator Factor incrementation
                        if (this._sarAccFactorLevel + this._sarAccFactorStep <= this._sarMaxAccFactorLevel)
                        {
                            this._accFactor.Add(this._sarAccFactorLevel + this._sarAccFactorStep);
                        }
                    }

                    // Get the SAR formula
                    this._sar = this._sarHistory[this._sarHistory.Count - 1] +
                                this._accFactor.Max() * (this._epData.Max() - this._sarHistory[this._sarHistory.Count - 1]);
                    this._sarHistory.Add(this._sar);
                }

            }

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
            if (this._sarHistory[this._sarHistory.Count - 1] < this._sarPricesHistory[this._sarPricesHistory.Count - 1] &&
                                         (!this._currentWay || // current way is sell
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
            else if (this._sarHistory[this._sarHistory.Count - 1] > this._sarPricesHistory[this._sarPricesHistory.Count - 1] &&
                                              (this._currentWay || // Current way is buy
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
    }
}
