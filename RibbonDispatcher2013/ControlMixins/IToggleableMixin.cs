using System;
using PGSolutions.RibbonDispatcher2013.AbstractCOM;

namespace PGSolutions.RibbonDispatcher2013.ControlMixins {
    /// <summary>The interface for controls that can be toggled.</summary>
    [CLSCompliant(true)]
    internal interface IToggleableMixin {
        /// <summary>TODO</summary>
        void OnChanged();

        /// <summary>TODO</summary>
        void OnToggled(bool IsPressed);
            
        /// <summary>TODO</summary>
        IRibbonTextLanguageControl LanguageStrings { get; }
    }
}
