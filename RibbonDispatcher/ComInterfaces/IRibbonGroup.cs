﻿////////////////////////////////////////////////////////////////////////////////////////////////////
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
    public interface IRibbonGroup {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        string Id           { get; }

        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        string Description  { get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [DispId(DispIds.KeyTip)]
        [Description("Returns the KeyTip string for this control.")]
        string KeyTip       { get; }

        /// <summary>Returns the Label string for this control.</summary>
        [DispId(DispIds.Label)]
        [Description("Returns the Label string for this control.")]
        string Label        { get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [DispId(DispIds.ScreenTip)]
        [Description("Returns the screenTip string for this control.")]
        string ScreenTip    { get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [DispId(DispIds.SuperTip)]
        [Description("Returns the SuperTip string for this control.")]
        string SuperTip     { get; }

        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [Description("Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.")]
        [DispId(DispIds.SetLanguageStrings)]
        void   SetLanguageStrings(IRibbonControlStrings languageStrings);

        /// <summary>.</summary>
        [Description(".")]
        [DispId(DispIds.IsEnabled)]
        bool   IsEnabled    { get; set; }
        /// <summary>.</summary>
        [Description(".")]
        //    [DispId(DispIds.IsVisible)]
        bool   IsVisible    { get; set; }

        /// <summary>.</summary>
        [Description(".")]
        void Invalidate();
    }

    public interface ICustomRibbonGroup : IRibbonGroup {
        /// <summary>Sets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Sets whether or not inactive controls should be visible on the Ribbon.")]
        void SetShowWhenInactive(bool showInactive);
    }
}
