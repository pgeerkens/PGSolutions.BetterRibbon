////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>Additional implementation-specific methods exposed by the Callback ModelFactory.</summary>
    public interface IRibbonViewModel {
        /// <summary>The Ribbon ControlID of the Ribbon definition being dispatched by this instance.</summary>
        string           ControlId        { get; }

        /// <summary>.</summary>
        ViewModelFactory ViewModelFactory { get; }

        /// <summary>.</summary>
        IRibbonUI        RibbonUI         { get; }
    }
}
