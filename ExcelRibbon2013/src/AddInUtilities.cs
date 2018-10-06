using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher;

namespace PGSolutions.ExcelRibbon2013 {
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAddInUtilities {
        void ImportData();

        IMain NewMain();
    }

    [Serializable]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IAddInUtilities))]
    public class AddInUtilities : IAddInUtilities {
        // This method tries to write a string to cell A1 in the active worksheet.
        public void ImportData() {
            var activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            if (activeWorksheet != null) {
                var range1 = activeWorksheet.get_Range("A1", Type.Missing);
                range1.Value2 = "This is my data";
            }
        }

        /// <inheritdoc/>
        [Description("Returns a new instance of the entry handle for {RibbonDispatcher}.")]
        public IMain NewMain() => new Main();
    }
}
