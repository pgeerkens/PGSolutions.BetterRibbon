////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The total interface (required to be) exposed externally by ButtonVM objects.</summary>
    [CLSCompliant(false)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonButton)]
    public interface IRibbonButton : IRibbonControlVM, IRibbonImageable {
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
        [Description("Gets or sets whether the control is enabled.")]
        new bool IsEnabled      { get; }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [Description("Gets or sets whether the control is visible.")]
        new bool IsVisible      { get; }

        /// <inheritdoc/>
        new void Invalidate();

        /// <summary>Gets or sets the preferred {RibbonControlSize} for the control.</summary>
        [Description("Gets or sets the preferred {RdControlSize} for the control.")]
        bool         IsLarge    { get; }

        /// <summary>Callback for the Clicked event on the control.</summary>
        [Description("Callback for the Clicked event on the control.")]
        void OnClicked(object sender, EventArgs e);

        /// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        [Description("Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.")]
        new ImageObject Image   { get; }
        /// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        [Description("Gets or sets whether to show the control's image; ignored by Large controls.")]
        new bool ShowImage      { get; }
        /// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        [Description("Gets or sets whether to show the control's label; ignored by Large controls.")]
        new bool ShowLabel      { get; }
    }
}
