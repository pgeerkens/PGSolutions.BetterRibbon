////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    [CLSCompliant(false)]
    public class VbaExportSelectedEventArgs : VbaExportEventArgs {
        public VbaExportSelectedEventArgs(ProjectFilter filter, FileDialogSelectedItems files) : base(filter)
        => Files = files;

        public FileDialogSelectedItems Files { get; }
    }

    [CLSCompliant(false)]
    public class VbaExportCurrentEventArgs : VbaExportEventArgs {
        public VbaExportCurrentEventArgs(ProjectFilter filter, Workbook workbook) : base(filter)
        => ActiveWorkbook = workbook;

        public Workbook ActiveWorkbook { get; }
    }

    [CLSCompliant(false)]
    public class VbaExportEventArgs : EventArgs {
        public VbaExportEventArgs(ProjectFilter filter) : base() => ProjectFilter = filter;

        public ProjectFilter ProjectFilter { get; }
    }
}
