////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses
{
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IControlStrings))]
    [Guid(Guids.ControlStrings)]
    public class ControlStrings : IControlStrings {
        public static ControlStrings Empty { get; } = new ControlStrings();
        private ControlStrings() : this(null) { }

        /// <summary>TODO</summary>
        [SuppressMessage( "Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage." )]
        public ControlStrings(
            string label,
            string screenTip      = null,
            string superTip       = null,
            string keyTip         = null,
            string alternateLabel = null,
            string description    = null
        ) {
            Label           = label;
            ScreenTip       = screenTip;
            SuperTip        = superTip;
            KeyTip          = keyTip;
            AlternateLabel  = alternateLabel;
            Description     = description;
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

        public static ControlStrings Default(string Id) => new ControlStrings(Id);
    }
}
