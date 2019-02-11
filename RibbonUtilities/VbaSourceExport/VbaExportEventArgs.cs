////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    [CLSCompliant(false)]
    public class VbaExportEventArgs : EventArgs {
        public VbaExportEventArgs(ProjectFilter filter) : this(filter,null) { }

        public VbaExportEventArgs(ProjectFilter filter, FileDialogSelectedItems files) : base() {
            ProjectFilter = filter;
            Files         = files;
        }

        public FileDialogSelectedItems Files         { get; }
        public ProjectFilter           ProjectFilter { get; }
    }
}
