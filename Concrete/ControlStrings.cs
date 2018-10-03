using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.AbstractCOM;

namespace PGSolutions.RibbonDispatcher.Concrete {

    /// <summary>A Dictionary-based ControlStrings implementation.</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IControlStrings))]
    [Guid(Guids.ControlStrings)]
    public class ControlStrings : IControlStrings {
        /// <summary>Creates a new empty ControlStrings collection.</summary>
        public ControlStrings() => _dictionary = new Dictionary<string,string>();

        Dictionary<string,string> _dictionary;

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object)")]
        public string AddControl(string ControlId,
            string Label          = null,
            string ScreenTip      = null,
            string SuperTip       = null,
            string AlternateLabel = null,
            string Description    = null,
            string KeyTip         = null
        ) {
            _dictionary.AddNotNull($"{ControlId}_Label",          Label);
            _dictionary.AddNotNull($"{ControlId}_ScreenTip",      ScreenTip);
            _dictionary.AddNotNull($"{ControlId}_SuperTip",       SuperTip);
            _dictionary.AddNotNull($"{ControlId}_AlternateLabel", AlternateLabel);
            _dictionary.AddNotNull($"{ControlId}_Description",    Description);
            _dictionary.AddNotNull($"{ControlId}_KeyTip",         KeyTip);
            return ControlId;
        }

        /// <inheritdoc/>
        public int         Count              => _dictionary.Count;
        /// <inheritdoc/>
        public string      this[string Index] => _dictionary.FirstOrDefault(i => i.Key == Index).Value
                                              ?? Index + " Unknown";
    }
}
