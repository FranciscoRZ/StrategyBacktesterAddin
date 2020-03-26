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
        private readonly FIFODoubleArray _sarHistory;

        /// <summary>
        /// FIFO collection with a Price history with Price(i-1) and Price(i)
        /// </summary>
        private readonly FIFODoubleArray _sarPricesHistory;

        /// <summary>
        /// FIFO collection with a all data to calculate the Extremum Point EP
        /// </summary>
        public List<double> epData = new List<double>();

        /// <summary>
        /// Extremum Point computed then stored in epData
        /// </summary>
        private double _ep;

        /// <summary>
        /// Accelerator factor (a) to use from initial a to max a
        /// </summary>
        public List<double> accFactor = new List<double>();

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
            this._sarHistory = new FIFODoubleArray(3);
            this._sarPricesHistory = new FIFODoubleArray(3);
        }

        /// <summary>
        /// Computes the indicator and steps through the strategy (Open or Close position)
        /// </summary>
        /// <param name="arrivedQuote">Used to update all the parameters of the trading strategy.</param>
        /// <returns>true if the position is opened (or flipped), false otherwise</returns>
        public override bool Step(Quote arrivedQuote)
        {
            // Update the data arrays
            this._sarPricesHistory.Put(arrivedQuote.ClosePrice);

            // Fisrt SAR Formula 
            if (this._sarHistory.GetSum() == 0.00) // Initialisation
            {
                this._sar = arrivedQuote.ClosePrice;
                this._sarHistory.Put(this._sar);
                this.accFactor.Add(this._sarAccFactorLevel);
            }

            // Fisrt calcul for the Extremum Point
            else if (this._sarHistory.Get(1) == 0)
            {
                this._ep = this._sarPricesHistory.Get(1);
                this.epData.Add(this._ep);

                // Get the SAR formula
                this._sar = this._sarHistory.Get(0) + this._sarAccFactorLevel * (this._ep - this._sarHistory.Get(0));
                this._sarHistory.Put(this._sar);
            }

            // Extremum Point calculation
            else
            {
                if (this._sarHistory.Get(1) < this._sarPricesHistory.Get(1))
                {

                    // Same trend
                    if (this._sarHistory.Get(0) <= this._sarPricesHistory.Get(0))
                    {
                        this._ep = Math.Max(this._sarPricesHistory.Get(2), this.epData.Max());
                        this.epData.Add(this._ep);
                    }

                    // Different trend
                    else
                    {
                        this._ep = this._sarPricesHistory.Get(2);
                        this.epData.Clear();
                        this.epData.Add(this._ep);
                        // Accelerator Factor incrementation
                        if (this._sarAccFactorLevel + this._sarAccFactorStep <= this._sarMaxAccFactorLevel)
                        {
                            this.accFactor.Add(this._sarAccFactorLevel + this._sarAccFactorStep);
                        }
                    }

                    // Get the SAR formula
                    this._sar = this._sarHistory.Get(0) + this.accFactor.Max() * (this.epData.Max() - this._sarHistory.Get(0));
                    this._sarHistory.Put(this._sar);
                }

                else
                {
                    // Same trend
                    if (this._sarHistory.Get(0) >= this._sarPricesHistory.Get(1))
                    {
                        this._ep = Math.Min(this._sarPricesHistory.Get(2), this.epData.Min());
                        this.epData.Add(this._ep);
                    }
                    // Different trend
                    else
                    {
                        this._ep = this._sarPricesHistory.Get(2);
                        this.epData.Clear();
                        this.epData.Add(this._ep);
                        // Accelerator Factor incrementation
                        if (this._sarAccFactorLevel + this._sarAccFactorStep <= this._sarMaxAccFactorLevel)
                        {
                            this.accFactor.Add(this._sarAccFactorLevel + this._sarAccFactorStep);
                        }
                    }

                    // Get the SAR formula
                    this._sar = this._sarHistory.Get(0) + this.accFactor.Max() * (this.epData.Max() - this._sarHistory.Get(0));
                    this._sarHistory.Put(this._sar);
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
            if (this._sarHistory.Get(2) <= this._sarPricesHistory.Get(2) && (!this._currentWay || // current way is sell
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
            else if (this._sarHistory.Get(2) > this._sarPricesHistory.Get(2) && (this._currentWay || // Current way is buy
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
