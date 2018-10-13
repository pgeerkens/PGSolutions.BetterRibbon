////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Windows.Forms;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal class ProjectFilterExcel : ProjectFilter  {
        public ProjectFilterExcel() : this("", "") { }

        public ProjectFilterExcel(string description, string extensions) : base(description, extensions) { }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            if ( IsProjectModelTrusted) {
                foreach (string selectedItem in items) {
                    ExtractProject(selectedItem, destIsSrc);
                    // DoEvents
                }
            } else {
                MessageBox.Show("Please enable trust of the Project Object Model", "Project Model Not Trusted",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        /// <summary>Returns true exactly when the Project Object Model is trusted.</summary>
        private static bool IsProjectModelTrusted => Globals.ThisAddIn.Application.VBE != null;

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private static void ExtractProject(string filename, bool destIsSrc) {
            Excel.Workbook wkbk = null;
            try {
                wkbk = Globals.ThisAddIn.Application.Workbooks.Open(filename, null, true);
                ExtractOpenProject(wkbk, destIsSrc);
            //} catch (IOException ex) {
            //    ExtractClosedProject(app, filename, destIsSrc);
            } finally {
                wkbk?.Close();
            }
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private static void ExtractClosedProject(Excel.Application app, string filename, bool destIsSrc) {
            var wkbk = app.Workbooks.Open(filename, UpdateLinks:false, ReadOnly:true, AddToMru:false, Editable:false);

            try {
                ExtractOpenProject(wkbk, destIsSrc);
            } finally {
                wkbk?.Close();
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        public static void ExtractOpenProject(Excel.Workbook wkbk, bool destIsSrc) =>
            ExtractProjectModules(wkbk.VBProject, CreateDirectory(wkbk.FullName, destIsSrc));
    }

}
