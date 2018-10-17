////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
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
    [ComDefaultInterface(typeof(IRibbonControlStrings))]
    [Guid(Guids.RibbonTextLanguageControl)]
    public class RibbonControlStrings : IRibbonControlStrings {
        public static RibbonControlStrings Empty { get; } = new RibbonControlStrings();
        private RibbonControlStrings() : this("", "", "", "", "", "") { }

        /// <summary>TODO</summary>
        [SuppressMessage( "Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage." )]
        public RibbonControlStrings(
            string label,
            string screenTip      = null,
            string superTip       = null,
            string keyTip         = null,
            string alternateLabel = null,
            string description    = null
        ) {
            Label           = label         ?? "Missing";
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

        public static RibbonControlStrings Default(string Id) => new RibbonControlStrings(Id);
    }
}
