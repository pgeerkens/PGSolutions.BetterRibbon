////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    internal interface IProjectFilter {
        /// <summary>Returns the Description for this filter.</summary>
        string Description { get; }

        /// <summary>Returns the Extensions list for this filter.</summary>
        string Extensions { get; }

        IApplication Application { get; }

        /// <summary>Exports modules from specified Access databases to eponymous subdirectories.</summary>
        void   ExtractProjects(FileDialogSelectedItems Items,bool destIsSrc);
    }
}
