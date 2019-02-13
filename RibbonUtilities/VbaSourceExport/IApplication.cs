////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VBE = Microsoft.Vbe.Interop;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public interface IApplication {
        /// <summary>.</summary>
        void DoOnOpenWorkbook(string wkbkFullName, Action<VBE.VBProject, string> action);

        /// <summary>.</summary>
        bool DisplayAlerts { get; set; }
    }
}
