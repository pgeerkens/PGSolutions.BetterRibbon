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
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        string Description { get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        new string KeyTip { get; }

        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        new string Label { get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        new string ScreenTip { get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        new string SuperTip { get; }
    }

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IControlStrings2))]
    [Guid(Guids.ControlStrings2)]
    public class ControlStrings2 : ControlStrings, IControlStrings2 {
        public new static ControlStrings2 Empty { get; } = new ControlStrings2();
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
        => Description     = description;
        /// <inheritdoc/>
        public string Description { get; }

        public new static ControlStrings2 Default(string Id) => new ControlStrings2(Id);
    }
}
