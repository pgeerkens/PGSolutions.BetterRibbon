////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The interface for the Ribbon ViewModelFactory.</summary>
    public interface IViewModelFactory {
        /// <summary>.</summary>
        /// <param name="controlId"></param>
        [Description("")]
        IControlStrings GetStrings(string controlId);

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "");

        object LoadImage(string imageId);
    }
}
