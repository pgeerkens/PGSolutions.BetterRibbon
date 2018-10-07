using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher2013.AbstractCOM {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonButton)]
    public interface IRibbonButton {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
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
        [Description("Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.")]
        void SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(DispIds.IsEnabled)]
        [Description("Gets or sets whether the control is enabled.")]
        bool IsEnabled          { get; set; }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(DispIds.IsVisible)]
        [Description("Gets or sets whether the control is visible.")]
        bool IsVisible          { get; set; }

        /// <summary>Gets or sets the preferred {RdControlSize} for the control.</summary>
        [DispId(DispIds.Size)]
        [Description("Gets or sets the preferred {RdControlSize} for the control.")]
        RdControlSize Size      { get; set; }

        /// <summary>Callback for the Clicked event on the control.</summary>
        [DispId(DispIds.OnClicked)]
        [Description("Callback for the Clicked event on the control.")]
        void OnClicked();

        /// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        [DispId(DispIds.Image)]
        [Description("Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.")]
        object Image            { get; }
        /// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        [DispId(DispIds.ShowImage)]
        [Description("Gets or sets whether to show the control's image; ignored by Large controls.")]
        bool ShowImage          { get; set; }
        /// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        [DispId(DispIds.ShowLabel)]
        [Description("Gets or sets whether to show the control's label; ignored by Large controls.")]
        bool ShowLabel          { get; set; }
        /// <summary>Sets the current Image for the control as an {IPictureDisp}.</summary>
        [DispId(DispIds.SetImageDisp)]
        [Description("Sets the current Image for the control as an {IPictureDisp}.")]
        void SetImageDisp(IPictureDisp Image);
        /// <summary>Sets the current Image for the control as a {string} naming an MsoImage.</summary>
        [DispId(DispIds.SetImageMso)]
        [Description("Sets the current Image for the control as a {string} naming an MsoImage.")]
        void SetImageMso(string ImageMso);
    }
}
