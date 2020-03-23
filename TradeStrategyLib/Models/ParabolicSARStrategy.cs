using System;
using System.Collections.Generic;

using DataTypes;
using System.Linq;

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
        private readonly double _SARAccFactorLevel;

        /// <summary>
        /// Maximum Accelerator Factor
        /// </summary>
        private readonly double _SARMaxAccFactorLevel;

        /// <summary>
        /// Accelerator Factor Step
        /// </summary>
        private readonly double _SARAccFactorStep;

        /// <summary>
        /// FIFO collection with a SAR history with SAR(i-1) and SAR(i)
        /// </summary>
        private readonly FIFODoubleArray _SARHistory;

        /// <summary>
        /// FIFO collection with a Price history with Price(i-1) and Price(i)
        /// </summary>
        private readonly FIFODoubleArray _SARPricesHistory;

        /// <summary>
        /// FIFO collection with a all data to calculate the Extremum Point EP
        /// </summary>
        private readonly FIFODoubleArray _EPdata;
        private double EP;

        /// <summary>
        /// Accelerator factor (a) to use from initial a to max a
        /// </summary>
        private readonly FIFODoubleArray _AccFactor;

        /// <summary>
        /// SAR Formula
        /// </summary>
        private double SAR;

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
            this._SARAccFactorLevel = SARAccFactor;
            this._SARMaxAccFactorLevel = SARMaxAccFactor;
            this._SARAccFactorStep = SARAccFactorStep;
            this._SARHistory = new FIFODoubleArray(2);
            this._SARPricesHistory = new FIFODoubleArray(3);
            this._EPdata = new FIFODoubleArray(1000);
            this._AccFactor = new FIFODoubleArray(2);// (int)Math.Ceiling(SARMaxAccFactor / SARAccFactorStep));
        }

        /// <summary>
        /// Computes the indicator and steps through the strategy (Open or Close position)
        /// </summary>
        /// <param name="arrivedQuote">Used to update all the parameters of the trading strategy.</param>
        /// <returns>true if the position is opened (or flipped), false otherwise</returns>
        public override bool Step(Quote arrivedQuote)
        {
            // Update the data arrays
            this._SARPricesHistory.Put(arrivedQuote.ClosePrice);

            // Fisrt SAR Formula 
            if (this._SARHistory.GetSum() == 0.00) // Initialisation
            {
                SAR = arrivedQuote.ClosePrice;
                this._SARHistory.Put(SAR);

                this._AccFactor.Put(this._SARAccFactorLevel);
            }

            // Fisrt calcul for the Extremum Point
            else if (this._SARHistory.Get(1) == 0)
            {
                EP = this._SARPricesHistory.Get(1);
                this._EPdata.Put(EP);

                // Get the SAR formula
                SAR = this._SARHistory.Get(0) + this._AccFactor.GetMax() * (EP - this._SARHistory.Get(0));
                this._SARHistory.Put(SAR);
            }

            // Extremum Point calculation
            else
            {
                if(this._SARHistory.Get(1) < this._SARPricesHistory.Get(1))
                {

                    // Same trend
                    if(this._SARHistory.Get(0) <= this._SARPricesHistory.Get(0))
                    {
                        EP = Math.Max(this._SARPricesHistory.Get(2), this._EPdata.GetMax());
                        this._EPdata.Put(EP);
                    }

                    // Different trend
                    else
                    {
                        EP = this._SARPricesHistory.Get(2);
                        this._EPdata.Initialize();
                        this._EPdata.Put(EP);
                        // Accelerator Factor incrementation
                        if (this._SARAccFactorLevel + this._SARAccFactorStep <= this._SARMaxAccFactorLevel)
                        {
                            this._AccFactor.Put(this._SARAccFactorLevel + this._SARAccFactorStep);
                        }
                    }

                    // Get the SAR formula
                    SAR = this._SARHistory.Get(0) + this._AccFactor.GetMax() * (this._EPdata.GetMax() - this._SARHistory.Get(0));
                    this._SARHistory.Put(SAR);
                }

                else
                {
                    // Same trend
                    if (this._SARHistory.Get(0) >= this._SARPricesHistory.Get(1))
                    {
                        EP = Math.Min(this._SARPricesHistory.Get(2), this._EPdata.GetMin());
                        this._EPdata.Put(EP);
                    }
                    // Different trend
                    else
                    {
                        EP = this._SARPricesHistory.Get(2);
                        this._EPdata.Initialize();
                        this._EPdata.Put(EP);
                        // Accelerator Factor incrementation
                        if (this._SARAccFactorLevel + this._SARAccFactorStep <= this._SARMaxAccFactorLevel)
                        {
                            this._AccFactor.Put(this._SARAccFactorLevel + this._SARAccFactorStep);
                        }
                    }

                    // Get the SAR formula
                    SAR = this._SARHistory.Get(0) + this._AccFactor.GetMax() * (this._EPdata.GetMax() - this._SARHistory.Get(0));
                    this._SARHistory.Put(SAR);
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
            if (this._SARHistory.Get(1) <= this._SARPricesHistory.Get(2) && (!this._currentWay || // current way is sell
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
            else if (this._SARHistory.Get(1) > this._SARPricesHistory.Get(2) && (this._currentWay|| // Current way is buy
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
