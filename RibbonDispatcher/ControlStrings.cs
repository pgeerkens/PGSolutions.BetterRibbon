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
        /// <summary>Returns the Label string for this control.</summary>
        [DispId(1)]
        string Label {
        [Description("Returns the Label string for this control.")]
            get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [DispId(2)]
        string ScreenTip {
        [Description("Returns the screenTip string for this control.")]
            get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [DispId(3)]
        string SuperTip {
        [Description("Returns the SuperTip string for this control.")]
            get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [DispId(4)]
        string KeyTip {
        [Description("Returns the KeyTip string for this control.")]
            get; }
    }

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IControlStrings))]
    [Guid(Guids.ControlStrings)]
    public class ControlStrings: IControlStrings {
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
