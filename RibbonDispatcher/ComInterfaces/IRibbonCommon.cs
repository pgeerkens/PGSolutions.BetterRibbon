////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The base interface for Ribbnon controls.</summary>
    [CLSCompliant(true)]
    public interface IRibbonCommon {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        string Id           { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        string Description  { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        string KeyTip       { get; }
        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        string Label        { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        string ScreenTip    { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        string SuperTip     { get; }

        /// <summary>Gets or sets whether or not the control is enabled.</summary>
        [Description("Gets or sets whether or not the control is enabled.")]
        bool IsEnabled      { get; }
        /// <summary>Gets or sets whether or not the control is visible.</summary>
        [Description("Gets or sets whether or not the control is visible..")]
        bool IsVisible      { get; }

        void Invalidate();

        void Detach();
    }
}
