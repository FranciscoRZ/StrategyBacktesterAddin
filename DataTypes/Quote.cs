using System;

namespace DataTypes
{
    public class Quote : ICloneable
    {
        /// <summary>
        /// Unique generated Hash Code
        /// </summary>
        private int generatedHashCodeConstant = new Random().Next();

        /// <summary>
        /// Gets or sets the TimeStamp of this quote
        /// </summary>
        public DateTime TimeStamp { get; set; }

        /// <summary>
        /// Get or sets the Open price of this quote
        /// </summary>
        public double OpenPrice { get; set; }

        /// <summary>
        /// Get or sets the High price of this quote
        /// </summary>
        public double HighPrice { get; set; }

        /// <summary>
        /// Gets or sets the Low price of this quote
        /// </summary>
        public double LowPrice { get; set; }

        /// <summary>
        /// Gets or sets the Close price of this quote
        /// </summary>
        public double ClosePrice { get; set; }

        /// <summary>
        /// Gets or sets the Volume of this quote
        /// </summary>
        public double Volume { get; set; }

        /// <summary>
        /// Clone this object and return a new instance of it
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            Quote returnedClone = new Quote
            {
                TimeStamp = new DateTime(this.TimeStamp.Ticks),
                OpenPrice = this.OpenPrice,
                HighPrice = this.HighPrice,
                LowPrice = this.LowPrice,
                ClosePrice = this.ClosePrice,
                Volume = this.Volume
            };
            return returnedClone;
        }

        /// <summary>
        /// Override object.Equals
        /// </summary>
        /// <param name="obj">Other object</param>
        /// <returns>True if its the same object</returns>
        public override bool Equals(object obj)
        {
            if (obj == null || this.GetType() != obj.GetType())
            {
                return false;
            }

            var other = (Quote)obj;
            if (this.TimeStamp != other.TimeStamp)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Override object.GetHashCode
        /// </summary>
        /// <returns>Hash code</returns>
        public override int GetHashCode()
        {
            return (int)this.TimeStamp.Ticks ^ this.generatedHashCodeConstant;
        }

    }
}
