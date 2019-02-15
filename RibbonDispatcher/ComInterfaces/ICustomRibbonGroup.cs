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
    [Guid(Guids.ICustomRibbonGroup)]
    public interface ICustomRibbonGroup : IRibbonGroup {
        /// <summary>Sets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Sets whether or not inactive controls should be visible on the Ribbon.")]
        void SetShowInactive(bool showInactive);

        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        new string Id { get; }

        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        new string Description { get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [DispId(DispIds.KeyTip)]
        [Description("Returns the KeyTip string for this control.")]
        new string KeyTip { get; }

        /// <summary>Returns the Label string for this control.</summary>
        [DispId(DispIds.Label)]
        [Description("Returns the Label string for this control.")]
        new string Label { get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [DispId(DispIds.ScreenTip)]
        [Description("Returns the screenTip string for this control.")]
        new string ScreenTip { get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [DispId(DispIds.SuperTip)]
        [Description("Returns the SuperTip string for this control.")]
        new string SuperTip { get; }

        ///// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        //[Description("Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.")]
        //[DispId(DispIds.SetLanguageStrings)]
        //new void SetLanguageStrings(IRibbonControlStrings strings);

        /// <summary>.</summary>
        [Description(".")]
        [DispId(DispIds.IsEnabled)]
        new bool IsEnabled { get; set; }
        /// <summary>.</summary>
        [Description(".")]
        //    [DispId(DispIds.IsVisible)]
        new bool IsVisible { get; set; }

        /// <summary>.</summary>
        [Description(".")]
        new void Invalidate();

        /// <summary>.</summary>
        [Description(".")]
        IRibbonCommon Attach();

        /// <summary>.</summary>
        [Description(".")]
        void DetachControls();
    }
}
