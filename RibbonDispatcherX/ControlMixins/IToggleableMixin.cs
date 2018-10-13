using System;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ControlMixins {
    /// <summary>The interface for controls that can be toggled.</summary>
    [CLSCompliant(true)]
    internal interface IToggleableMixin {
        /// <summary>TODO</summary>
        void OnChanged();

        /// <summary>TODO</summary>
        void OnToggled(bool IsPressed);
            
        /// <summary>TODO</summary>
        IRibbonControlStrings LanguageStrings { get; }
    }
}
