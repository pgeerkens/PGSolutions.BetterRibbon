////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher {
    /// <summary>The full interface required from string-suppliers, for controls supporting Description.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IControlStrings2)]
    public interface IControlStrings2 : IControlStrings {
        /// <summary>Returns the Label string for this control.</summary>
        [DispId(1)]
        new string Label {
        [Description("Returns the Label string for this control.")]
            get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [DispId(2)]
        new string ScreenTip {
        [Description("Returns the screenTip string for this control.")]
            get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [DispId(3)]
        new string SuperTip {
        [Description("Returns the SuperTip string for this control.")]
            get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [DispId(4)]
        new string KeyTip {
        [Description("Returns the KeyTip string for this control.")]
            get; }

        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [DispId(5)]
        string Description {
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
            get; }
    }

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IControlStrings2))]
    [Guid(Guids.ControlStrings2)]
    public class ControlStrings2 : ControlStrings, IControlStrings2 {
        private ControlStrings2() : this(null) { }

        /// <summary>TODO</summary>
        [SuppressMessage( "Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage." )]
        internal ControlStrings2(
            string label,
            string screenTip      = null,
            string superTip       = null,
            string keyTip         = null,
            string description    = null
        ) : base(label, screenTip, superTip,keyTip)
        => Description = description;
        /// <inheritdoc/>
        public string Description { get; }

        public new static ControlStrings2 Default(string Id) => new ControlStrings2(Id);
    }
}
