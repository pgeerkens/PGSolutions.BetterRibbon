////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
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
        public ControlStrings(
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
        public ControlStrings2(
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
