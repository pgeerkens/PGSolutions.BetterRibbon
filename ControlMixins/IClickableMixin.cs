using System;

namespace PGSolutions.RibbonDispatcher.ControlMixins {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    [CLSCompliant(true)]
    public interface IClickableMixin {
        /// <summary>TODO</summary>
        void OnClicked();
    }
}
