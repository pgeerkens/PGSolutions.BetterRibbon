////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonToggleButton)]
    public interface IRibbonToggleButton : IRibbonImageable {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        string Id               { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        string Description      { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        [DispId(DispIds.KeyTip)]
        [Description("Returns the KeyTip string for this control.")]
        string KeyTip           { get; }
        /// <summary>Returns the Label string for this control.</summary>
        [DispId(DispIds.Label)]
        [Description("Returns the Label string for this control.")]
        string Label            { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        [DispId(DispIds.ScreenTip)]
        [Description("Returns the screenTip string for this control.")]
        string ScreenTip        { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        [DispId(DispIds.SuperTip)]
        [Description("Returns the SuperTip string for this control.")]
        string SuperTip         { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId(DispIds.SetLanguageStrings)]
        void          SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(DispIds.IsEnabled)]
        [Description("Gets or sets whether the control is enabled.")]
        bool IsEnabled          { get; set; }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(DispIds.IsVisible)]
        [Description("Gets or sets whether the control is visible.")]
        bool IsVisible          { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.Size)]
        RdControlSize Size      { get; set; }

        /// <summary>Returns whether the control is pressed.</summary>
        [DispId(DispIds.IsPressed)]
        [Description("Returns whether the control is pressed.")]
        bool IsPressed          { get; /*set;*/ }
        /// <summary>Callback for the Pressed event on the control.</summary>
        [DispId(DispIds.OnToggled)]
        [Description("Callback for the Pressed event on the control.")]
        void OnToggled(bool IsPressed);

        /// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        [DispId(DispIds.Image)]
        [Description("Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.")]
        new object Image        { get; }
        /// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        [DispId(DispIds.ShowImage)]
        [Description("Gets or sets whether to show the control's image; ignored by Large controls.")]
        new bool ShowImage      { get; set; }
        /// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        [DispId(DispIds.ShowLabel)]
        [Description("Gets or sets whether to show the control's label; ignored by Large controls.")]
        new bool ShowLabel      { get; set; }
        /// <summary>Sets the current Image for the control as an {IPictureDisp}.</summary>
        [DispId(DispIds.SetImageDisp)]
        [Description("Sets the current Image for the control as an {IPictureDisp}.")]
        new void SetImageDisp(IPictureDisp Image);
        /// <summary>Sets the current Image for the control as a {string} naming an MsoImage.</summary>
        [DispId(DispIds.SetImageMso)]
        [Description("Sets the current Image for the control as a {string} naming an MsoImage.")]
        new void SetImageMso(string ImageMso);
    }
}
