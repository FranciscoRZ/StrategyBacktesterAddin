using System;

namespace TradeStrategyLib.Models
{
    /// <summary>
    /// Implemenentation of First In First Out collection
    /// </summary>
    public class FIFODoubleArray
    {
        /// <summary>
        /// Used to limit calls to the object's properties
        /// array.Length - 1
        /// </summary>
        private readonly int _arrayLastIndex;

        /// <summary>
        /// Array storing all current values taken into account
        /// in the trade strategy.
        /// </summary>
        private readonly double[] _array;

        /// <summary>
        /// The index of the last inserted value into the array
        /// </summary>
        private int _arrayLastInsertedIndex = 0;

        /// <summary>
        /// Constructor of the class <see cref="FIFODoubleArray"/> class.
        /// </summary>
        /// <param name="arraySize">Size (base 1) of the array.</param>
        public FIFODoubleArray(int arraySize)
        {
            this._array = new double[arraySize];
            this._arrayLastIndex = arraySize - 1;
        }

        /// <summary>
        /// Inserts a value into the array overwriting the old one.
        /// </summary>
        /// <param name="value"></param>
        public void Put(double value)
        {
            this._array[this._arrayLastInsertedIndex++] = value;
            if (this._arrayLastInsertedIndex == this._arrayLastIndex + 1)
            {
                this._arrayLastInsertedIndex = 0;
            }
        }

        /// <summary>
        /// Returns a new array with values sorted chronologically 
        /// (up to countOfValues values)
        /// </summary>
        /// <param name="countOfValues">Total number of observations to return.</param>
        /// <returns>Collection of clones sized as requested.</returns>
        public double[] GetFirstNValues(int countOfValues)
        {
            double[] valuesArray = new double[countOfValues];

            // Start with the oldest observation
            int arrayLastIndMoving = this._arrayLastInsertedIndex;
            for (int i = 0; i < countOfValues; i++)
            {
                if (arrayLastIndMoving + i == this._arrayLastIndex)
                {
                    arrayLastIndMoving = 0;
                }
                valuesArray[i] = this._array[arrayLastIndMoving + i];
            }

            return valuesArray;
        }

        /// <summary>
        /// Gets the ith value of the array. 
        /// </summary>
        /// <param name="index">The absolute index of the requested value</param>
        /// <returns>Value stored at the requested index</returns>
        public double Get(int index)
        {
            return this._array[index];
        }

        /// <summary>
        /// Returns the pointer to the this._array
        /// </summary>
        /// <returns>Pointer to the actual collection of values</returns>
        public double[] GetArray()
        {
            return this._array;
        }

        /// <summary>
        /// Returns the sum of the array
        /// </summary>
        /// <returns>Array's sum</returns>
        public double GetSum()
        {
            double sum = 0.00;
            for (int i = 0; i <= this._arrayLastIndex; i++)
            {
                sum += this._array[i];
            }

            return sum;
        }

        /// <summary>
        /// Returns the average of the array
        /// </summary>
        /// <returns>Array's mean</returns>
        public double GetArrayMean()
        {
            return this.GetSum() / (this._arrayLastIndex + 1);
        }

        /// <summary>
        /// Returns the array's volatility, defined here as the 
        /// standard deviation of its values
        /// </summary>
        /// <returns>Array's volatility</returns>
        public double GetArrayVol()
        {
            double arrayMean = this.GetArrayMean();
            double sum = 0.00;
            for (int i = 0; i <= this._arrayLastIndex; i++)
            {
                sum += Math.Pow(this._array[i] - arrayMean, 2);
            }
            sum /= this._arrayLastIndex;

            return Math.Sqrt(sum);
        }

        /// <summary>
        /// Returns the array's length
        /// </summary>
        /// <returns>Array's length</returns>
        public int GetSize()
        {
            return this._array.Length;
        }

    }
}
