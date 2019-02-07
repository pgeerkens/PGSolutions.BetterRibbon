////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    [CLSCompliant(false)]
    public sealed class VbaSourceExportModel : IBooleanSource {
        public VbaSourceExportModel(IList<IVbaSourceExportViewModel> viewModels) {
            DestIsSrc   = false;
            ViewModels  = viewModels;
            foreach (var viewModel in ViewModels) {
                viewModel.SelectedProjectsClicked += ExportSelectedProjects;
                viewModel.CurrentProjectClicked   += ExportCurrentProject;
                viewModel.UseSrcFolderToggled     += UseSrcFolderToggled;
                viewModel.Attach(this);
            }
        }

        bool IBooleanSource.Getter() => DestIsSrc;

        private IApplication                     Application { get; }

        /// <summary>Fakse => file destination is eponymous directory; else directory named "SRC".</summary>
        private bool                             DestIsSrc   { get; set; }
        private IList<IVbaSourceExportViewModel> ViewModels  { get; }

        private void UseSrcFolderToggled(object sender, bool isPressed) {
            DestIsSrc = isPressed;
            foreach (var viewModel in ViewModels) {
                viewModel.SelectedProjectButton.IsEnabled = ! DestIsSrc;
                viewModel.Invalidate();
            }
        }

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportCurrentProject(object sender, VbaExportEventArgs e)
        => (e.ProjectFilter as ProjectFilterExcel)?.ExtractOpenProject(DestIsSrc);

        /// <summary>Extracts VBA modules from a selected EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportSelectedProjects(object sender, VbaExportEventArgs e)
        => e.ProjectFilter.ExtractProjects(e.Files, DestIsSrc);
    }
}
