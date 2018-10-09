////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
    public interface IRibbonImageable {
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
