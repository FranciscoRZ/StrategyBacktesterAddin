using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace StrategyBacktesterAddin
{
    /// <summary>
    /// Singleton thread safe pattern class in order to make sure no multiple references
    /// are made to the Excel application
    /// </summary>
    public sealed class XLSingleton
    {
        /// <summary>
        /// The Excel Application
        /// </summary>
        public Excel.Application XLApp = (Excel.Application)ExcelDnaUtil.Application;
        /// <summary>
        /// The Excel workbook from and to which we'll write.
        /// </summary>
        public Excel.Workbook XLWorkbook;

        // Instance to be used throughout application
        private static readonly XLSingleton _instance = new XLSingleton();

        /// <summary>
        /// Private constructor
        /// </summary>
        private XLSingleton()
        {
            this.XLWorkbook = this.XLApp.ActiveWorkbook;
        }

        static XLSingleton()
        {
        }

        /// <summary>
        /// Getter for the instance of our class.
        /// It returns the instance created on the first call to the method.
        /// </summary>
        public static XLSingleton Instance { get { return _instance; } }

    }
}
