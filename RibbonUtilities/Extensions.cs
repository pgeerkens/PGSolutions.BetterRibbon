////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;

    public static partial class Extensions {
        /// <summary>Returns the speccified <see cref="Workbook"/> exactly when it is already open.</summary>
        /// <param name="excel">The running instnce of Excel.</param>
        /// <param name="path">The absolute full=path and -name for the desired workbook.</param>
        internal static Excel.Workbook TryItem(this Excel.Workbooks workbooks, string fullName) {
            foreach (Excel.Workbook wb in workbooks) if (wb.FullName == fullName) return wb;
            return null;
        }

        internal static void InvokeWithShiftKey(this System.Action action) {
            const byte VK_LSHIFT = 0xA0;  // left shift key
            try {
                VK_LSHIFT.KeyDown();
                action();
            }
            finally {
                VK_LSHIFT.KeyUp();
            }
        }
    }
}
