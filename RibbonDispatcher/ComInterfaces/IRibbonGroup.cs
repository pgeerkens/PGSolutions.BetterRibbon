////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonGroup)]
    public interface IRibbonGroup : IRibbonCommon {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        new string Id           { get; }

        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        new string Description  { get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [DispId(DispIds.KeyTip)]
        [Description("Returns the KeyTip string for this control.")]
        new string KeyTip       { get; }

        /// <summary>Returns the Label string for this control.</summary>
        [DispId(DispIds.Label)]
        [Description("Returns the Label string for this control.")]
        new string Label        { get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [DispId(DispIds.ScreenTip)]
        [Description("Returns the screenTip string for this control.")]
        new string ScreenTip    { get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [DispId(DispIds.SuperTip)]
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
