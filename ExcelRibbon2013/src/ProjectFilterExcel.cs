////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace PGSolutions.ExcelRibbon2013 {
    internal partial class ProjectFilterExcel : ProjectFilter  {
        public ProjectFilterExcel() : this("", "") { }

        public ProjectFilterExcel(string description, string extensions) : base(description, extensions) { }

        /// <inheritdoc/>
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            if ( IsProjectModelTrusted) {
                foreach (string selectedItem in items) {
                    ExtractProject(Globals.ThisAddIn.Application, selectedItem, destIsSrc);
                    // DoEvents
                }
            } else {
                MessageBox.Show("Please enable trust of the Project Object Model", "Project Model Not Trusted",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>Returns true exactly when the Project Object Model is trusted.</summary>
        private bool IsProjectModelTrusted => Globals.ThisAddIn.Application.VBE != null;

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(Excel.Application app, string filename, bool destIsSrc) {
            Workbook wkbk = null;
            try {
                wkbk = Globals.ThisAddIn.Application.Workbooks.Open(filename, null, true);
                ExtractOpenProject(wkbk, destIsSrc);
            //} catch (IOException ex) {
            //    ExtractClosedProject(app, filename, destIsSrc);
            } finally {
                wkbk?.Close();
            }
        }

        private void ExtractClosedProject(Excel.Application app, string filename, bool destIsSrc) {
            var wkbk = app.Workbooks.Open(filename, UpdateLinks:false, ReadOnly:true, AddToMru:false, Editable:false);

            try {
                ExtractOpenProject(wkbk, destIsSrc);
            } finally {
                wkbk?.Close();
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        public static void ExtractOpenProject(Workbook wkbk, bool destIsSrc) =>
            ExtractModulesByProject(wkbk.VBProject, CreateDirectory(wkbk.FullName, destIsSrc));
    }

}
