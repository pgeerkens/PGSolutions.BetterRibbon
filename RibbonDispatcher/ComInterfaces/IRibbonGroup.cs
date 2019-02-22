////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    public interface IRibbonGroup : IRibbonControlVM {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        new string Id           { get; }

        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        new string Description  { get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        new string KeyTip       { get; }

        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        new string Label        { get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        new string ScreenTip    { get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        new string SuperTip     { get; }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        new bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; }

        /// <summary>Gets or sets whether the control is visible.</summary>
        new bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; }

        /// <summary>.</summary>
        [Description(".")]
        new void Invalidate();
    }
}
