////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.ExcelRibbon {
    internal interface IProjectFilter {
        /// <summary>Returns the Description for this filter.</summary>
        string Description { get; }
        /// <summary>Returns the Extensions list for this filter.</summary>
        string Extensions { get; }

        /// <summary>Exports modules from specified Access databases to eponymous subdirectories.</summary>
        void   ExtractProjects(FileDialogSelectedItems Items,bool destIsSrc);
    }
}
