////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.LinksAnalyzer;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>THe main interface for VBA to access the Ribbon dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonDispatcher)]
    public interface IRibbonDispatcher {
        /// <summary>TODO</summary>
        [Description( "" )]
        //[DispId(1)]
        void InvalidateControl(string ControlId);

        /// <summary>TODO</summary>
        [Description( "" )]
        //[DispId(2)]
        void Invalidate();

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed",
                Justification = "Matches COM usage.")]
        //[DispId(3)]
        IRibbonControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "");

        /// <summary>Deactivate the specified control, detaching any attached data source.</summary>
        /// <param name="controlId">The ID of the control to be detached.</param>
        [Description("Deactivate the specified control, detaching any attached data source.")]
        //[DispId(4)]
        void DetachProxy(string controlId);

        /// <summary>Sets ehether or not inactive controls should be visible on the Ribbon.</summary>
        /// <param name="showWhenInactive"></param>
        [Description("Sets ehether or not inactive controls should be visible on the Ribbon.")]
        //[DispId(5)]
        void ShowInactive(bool showWhenInactive);

        /// <summary>Attaches and activates the specified Button control.</summary>
        /// <param name="controlId">The </param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches and activates the specified Button control.")]
        //[DispId(6)]
        IRibbonButton AttachButton(string controlId, IRibbonControlStrings strings);

        /// <summary>Attaches an {IBooleanSource} to the specified ToggleButton control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches an {IBooleanSource} to the specified ToggleButton control.")]
        //[DispId(7)]
        IRibbonToggle AttachToggle(string controlId, IRibbonControlStrings strings,
                IBooleanSource source);

        /// <summary>Attaches an {IBooleanSource} to the specified CheckBox control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches an {IBooleanSource} to the specified CheckBox control.")]
        //[DispId(8)]
        IRibbonToggle AttachCheckBox(string controlId, IRibbonControlStrings strings,
                IBooleanSource source);

        /// <summary>Attaches an {IIntegerSource} to the specified DropDown control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches an {IIntegerSource} to the specified DropDown control.")]
        //[DispId(9)]
        IRibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings,
                IIntegerSource source);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringLoader {
        IRibbonControlStrings GetStrings(string ControlId);
    }

    [CLSCompliant(false)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IImageLoader {
        IPictureDisp GetImage(string ControlId);
    }
}
