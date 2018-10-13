////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using Access = Microsoft.Office.Interop.Access;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal class ProjectFilterAccess : ProjectFilter {
        public ProjectFilterAccess(string description, string extensions) : base(description, extensions) { }

        /// <summary>Exports modules from specified Access databases to eponymous subdirectories.</summary>
        /// <remarks>
        /// </remarks>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            using (var app = AccessWrapper.New()) {
                if (!AccessWrapper.IsAccessSupported) { throw new NotSupportedException("MS-Access not available on this machine."); }

                if (app.IsProjectModelTrusted) {
                    foreach (string item in items) {
                        ExtractProject(app, item, destIsSrc);
                        // DoEvents
                    }
                } else {
                    MessageBox.Show("Please enable trust of the Project Object Model", "Project Model Not Trusted",
                            MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private static void ExtractProject(AccessWrapper app, string filename, bool destIsSrc) {
            try {
                app.OpenDbWithuotAutoexec(filename);
                ExtractOpenProject(app, destIsSrc);
            //} catch (IOException ex) {
            //    ExtractClosedProject(app, filename, destIsSrc);
            } finally {
               app.CloseCurrentDb();
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private static void ExtractOpenProject(AccessWrapper app, bool destIsSrc) {
            var exportPath = CreateDirectory(app.CurrentProjectName, destIsSrc);
            ExportDaoDatabase(app.AccessApp, exportPath);
        }

        private const int    dbSqlPassThrough = 112;
        private const string indent           = ",\n    ";

        private static void ExportDaoDatabase(Access.Application app, string exportPath) => ExportDaoDatabase(app, exportPath, true);
        private static void ExportDaoDatabase(Access.Application app, string exportPath, bool queriesAsSql) {
            if (queriesAsSql) {
                foreach (QueryDef qd in app.CurrentDb().QueryDefs) {
                    var sql = qd.Type == dbSqlPassThrough ? qd.SQL
                                                          : qd.SQL.Replace(", ", indent);
                    File.WriteAllText(FullPath(exportPath, qd.Name, "sql"), sql);
                }
            } else {
                foreach (AccessObject ao in app.CurrentData.AllQueries) {
                    app.SaveAsText(AcObjectType.acQuery, ao.FullName, FullPath(exportPath, ao.FullName, "sql"));
                }
            }

            foreach (AccessObject ao in app.CurrentProject.AllForms) {
                app.SaveAsText(AcObjectType.acForm, ao.FullName, FullPath(exportPath, ao.FullName, "mac"));
            }

            foreach (AccessObject ao in app.CurrentProject.AllMacros) {
                app.SaveAsText(AcObjectType.acMacro, ao.FullName, FullPath(exportPath, ao.FullName, "form"));
            }

            foreach (AccessObject ao in app.CurrentProject.AllReports) {
                app.SaveAsText(AcObjectType.acMacro, ao.FullName, FullPath(exportPath, ao.FullName, "report"));
            }
        }

        private static string FullPath(string folder, string filename, string extension) =>
            Path.Combine(folder, Path.ChangeExtension(filename, extension));
    }
}
