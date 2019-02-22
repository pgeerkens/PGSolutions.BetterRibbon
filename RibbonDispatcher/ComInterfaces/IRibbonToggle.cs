////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The ViewModel interface exposed by Ribbon ToggleButtons and CheckBoxes.</summary>
    public interface IRibbonToggle : IRibbonControlVM, IRibbonImageable {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        new string Id        { get; }

        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        new string Description { get; }

        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        new string KeyTip    { get; }

        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        new string Label     { get; }

        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        new string ScreenTip { get; }

        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        new string SuperTip  { get; }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [Description("Gets or sets whether the control is enabled.")]
        new bool IsEnabled   { get; }
        /// <summary>Gets or sets whether the control is visible.</summary>
    //    [DispId(DispIds.IsVisible)]
        [Description("Gets or sets whether the control is visible.")]
        new bool IsVisible   { get; }

        /// <inheritdoc/>
        new void Invalidate();

        /// <summary>Returns whether the control is pressed.</summary>
        [Description("Returns whether the control is pressed.")]
        bool IsPressed       { get; }
        /// <summary>Callback for the Pressed event on the control.</summary>
        [Description("Callback for the Pressed event on the control.")]
        void OnToggled(object sender, bool IsPressed);

        /// <summary>TODO</summary>
        bool     IsSizeable  { get; }
        /// <summary>TODO</summary>
        bool     IsLarge     { get; }

        /// <summary>TODO</summary>
        bool     IsImageable { get; }
        /// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        [Description("Gets or sets whether to show the control's image; ignored by Large controls.")]
        new bool ShowImage   { get; }
        /// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        [Description("Gets or sets whether to show the control's label; ignored by Large controls.")]
        new bool ShowLabel   { get; }
        /// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        [Description("Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.")]
        new ImageObject Image { get; }
    }
}
