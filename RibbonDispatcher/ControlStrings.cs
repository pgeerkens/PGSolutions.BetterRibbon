////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher {
    /// <summary>The abridged interface required from string-suppliers.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IControlStrings)]
    public interface IControlStrings {
        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        string KeyTip           { get; }

        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        string Label            { get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        string ScreenTip        { get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        string SuperTip         { get; }
    }

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IControlStrings))]
    [Guid(Guids.ControlStrings)]
    public class ControlStrings: IControlStrings {
        public static ControlStrings Empty { get; } = new ControlStrings();
        private ControlStrings() : this(null) { }

        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        internal ControlStrings(
            string label,
            string screenTip = null,
            string superTip = null,
            string keyTip = null
        ) {
            Label     = label;
            ScreenTip = screenTip;
            SuperTip  = superTip;
            KeyTip    = keyTip;
        }
        /// <inheritdoc/>
        public string Label { get; }

        /// <inheritdoc/>
        public string ScreenTip { get; }

        /// <inheritdoc/>
        public string SuperTip { get; }

        /// <inheritdoc/>
        public string KeyTip { get; }

        public static ControlStrings Default(string Id) => new ControlStrings(Id);
    }
}
