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
        public Excel.Application XLApp = (Excel.Application)ExcelDnaUtil.Application;
        public Excel.Workbook XLWorkbook;

        /// <summary>
        /// Private constructor
        /// </summary>
        private XLSingleton()
        {
            this.XLWorkbook = this.XLApp.ActiveWorkbook;
        }

        public static XLSingleton Instance { get { return NestedXLSingleton.instance; } }

        private class NestedXLSingleton
        {
            /// <summary>
            /// Explicit static constructor to tell C# not to mark type as beforefieldinit
            /// </summary>
            static NestedXLSingleton()
            {
            }

            /// <summary>
            /// Instanciation happens on first call to instance, and never again.
            /// </summary>
            internal static readonly XLSingleton instance = new XLSingleton();
        }

    }
}
