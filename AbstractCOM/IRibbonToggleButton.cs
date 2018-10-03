using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonToggleButton)]
    public interface IRibbonToggleButton {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        string        Id            { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [DispId(DispIds.Description)]
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        string Description { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        [DispId(DispIds.KeyTip)]
        [Description("Returns the KeyTip string for this control.")]
        string KeyTip { get; }
        /// <summary>Returns the Label string for this control.</summary>
        [DispId(DispIds.Label)]
        [Description("Returns the Label string for this control.")]
        string Label { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        [DispId(DispIds.ScreenTip)]
        [Description("Returns the screenTip string for this control.")]
        string ScreenTip { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        [DispId(DispIds.SuperTip)]
        [Description("Returns the SuperTip string for this control.")]
        string SuperTip { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [DispId(DispIds.SetLanguageStrings)]
        void          SetLanguageStrings(IRibbonTextLanguageControl languageStrings);

        /// <summary>TODO</summary>
        [DispId(DispIds.IsEnabled)]
        bool          IsEnabled     { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.IsVisible)]
        bool          IsVisible     { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.Size)]
        RdControlSize Size          { get; set; }

        /// <summary>TODO</summary>
        [DispId(DispIds.IsPressed)]
        bool          IsPressed     { get; }
        /// <summary>TODO</summary>
        [DispId(DispIds.OnToggled)]
        void          OnToggled(bool IsPressed);

        /// <summary>TODO</summary>
        [DispId(DispIds.Image)]
        object        Image         { get; }
        /// <summary>Returns or set whether to show the control's image; ignored by Large controls.</summary>
        [DispId(DispIds.ShowImage)]
        bool          ShowImage     { get; set; }
        /// <summary>Returns or set whether to show the control's label; ignored by Large controls.</summary>
        [DispId(DispIds.ShowLabel)]
        bool          ShowLabel     { get; set; }
        /// <summary>TODO</summary>
        [DispId(DispIds.SetImageDisp)]
        void          SetImageDisp(IPictureDisp Image);
        /// <summary>TODO</summary>
        [DispId(DispIds.SetImageMso)]
        void          SetImageMso(string ImageMso);
    }
}
