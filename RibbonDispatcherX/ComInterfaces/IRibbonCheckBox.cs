////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonCheckBox)]
    public interface IRibbonCheckBox : IRibbonCommon {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
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
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId(DispIds.SetLanguageStrings)]
        new void SetLanguageStrings(IRibbonControlStrings languageStrings);

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(DispIds.IsEnabled)]
        [Description("Gets or sets whether the control is enabled.")]
        new bool IsEnabled      { get; set; }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(DispIds.IsVisible)]
        [Description("Gets or sets whether the control is visible.")]
        new bool IsVisible      { get; set; }

        /// <inheritdoc/>
        new void Invalidate();

        /// <summary>Returns whether the control is pressed.</summary>
        [DispId(DispIds.IsPressed)]
        [Description("Returns whether the control is pressed.")]
        bool IsPressed          { get; /*set;*/ }
        /// <summary>Callback for the Pressed event on the control.</summary>
        [DispId(DispIds.OnToggled)]
        [Description("Callback for the Pressed event on the control.")]
        void OnToggled(bool IsPressed);
    }
}
