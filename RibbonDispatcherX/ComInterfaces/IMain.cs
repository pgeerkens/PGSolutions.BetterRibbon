////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>THe main interface for VBA to access the Ribbon dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMain)]
    public interface IMain {
        /// <summary>TODO</summary>
        [Description("")]
        IRibbonFactory RibbonFactory { get; }

        /// <summary>Attaches and activates the specified Button control.</summary>
        /// <param name="controlId">The </param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches and activates the specified Button control.")]
        IRibbonButton AttachButton(string controlId, IRibbonControlStrings strings);

        /// <summary>Attaches an {IBooleanSource} to the specified ToggleButton control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches an {IBooleanSource} to the specified ToggleButton control.")]
        IRibbonToggleButton AttachToggle(string controlId, IRibbonControlStrings strings,
                IBooleanSource source);

        /// <summary>Attaches an {IBooleanSource} to the specified CheckBox control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches an {IBooleanSource} to the specified CheckBox control.")]
        IRibbonCheckBox AttachCheckBox(string controlId, IRibbonControlStrings strings,
                IBooleanSource source);

        /// <summary>Attaches an {IIntegerSource} to the specified DropDown control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description("Attaches an {IIntegerSource} to the specified DropDown control.")]
        IRibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings,
                IIntegerSource source);

        /// <summary>Sets ehether or not inactive controls should be visible on the Ribbon.</summary>
        /// <param name="showWhenInactive"></param>
        [Description("Sets ehether or not inactive controls should be visible on the Ribbon.")]
        void ShowInactive(bool showWhenInactive);

        /// <summary>Deactivate the specified control, detaching any attached data source.</summary>
        /// <param name="controlId">The ID of the control to be detached.</param>
        [Description("Deactivate the specified control, detaching any attached data source.")]
        void DetachProxy(string controlId);

        /// <summary>TODO</summary>
        [Description("")]
        /// <inheritdoc/>
        void InvalidateControl(string ControlId);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringLoader
    {
        IRibbonControlStrings GetStrings(string ControlId);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IImageLoader
    {
        IPictureDisp GetImage(string ControlId);
    }
}
