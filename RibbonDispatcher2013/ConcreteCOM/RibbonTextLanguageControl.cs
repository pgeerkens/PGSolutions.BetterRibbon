using System;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher2013.AbstractCOM;

namespace PGSolutions.RibbonDispatcher2013.ConcreteCOM
{
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonTextLanguageControl))]
    [Guid(Guids.RibbonTextLanguageControl)]
    public class RibbonTextLanguageControl : IRibbonTextLanguageControl {
        /// <summary>TODO</summary>
        public RibbonTextLanguageControl(
            string label,
            string screenTip,
            string superTip,
            string keyTip,
            string alternateLabel,
            string description
        ) {
            Label           = label         ?? throw new ArgumentNullException(nameof(label)); 
            ScreenTip       = screenTip     ?? Label; 
            SuperTip        = superTip      ?? "SuperTip text for " + Label; 
            KeyTip          = keyTip        ?? "";
            AlternateLabel  = alternateLabel?? Label; 
            Description     = description   ?? "Description for " + Label;
        }
        /// <inheritdoc/>
        public string Label { get; }

        /// <inheritdoc/>
        public string ScreenTip { get; }

        /// <inheritdoc/>
        public string SuperTip { get; }

        /// <inheritdoc/>
        public string KeyTip { get; }

        /// <inheritdoc/>
        public string AlternateLabel { get; }

        /// <inheritdoc/>
        public string Description { get; }
    }
}
