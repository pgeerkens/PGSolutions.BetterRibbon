////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonGroupModel)]
    public interface IRibbonGroupModel : IRibbonControlModel {
        /// <summary>Gets the {IRibbonControlStrings} for this control.</summary>
        new IRibbonControlStrings Strings {
            [Description("Gets the {IRibbonControlStrings} for this control.")]
            get;
        }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        new bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        new bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }

        /// <summary>Gets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Gets whether or not inactive controls should be visible on the Ribbon.")]
        bool ShowInactive { get; }

        /// <summary>Sets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Sets whether or not inactive controls should be visible on the Ribbon.")]
        void SetShowInactive(bool showInactive);

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IRibbonGroupModel Attach(string controlId);

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        new void Invalidate();

        ///// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        //[Description("Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.")]
        //[DispId(DispIds.SetLanguageStrings)]
        //void SetLanguageStrings(IRibbonControlStrings languageStrings);
    }

    public interface IRibbonControlModel {
        /// <summary>Gets the {IRibbonControlStrings} for this control.</summary>
        IRibbonControlStrings Strings   {
            [Description("Gets the {IRibbonControlStrings} for this control.")]
            get;
        }

         /// <summary>Gets or sets whether the control is enabled.</summary>
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }

        /// <summary>Gets or sets whether the control is visible.</summary>
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }

        void Invalidate();
    }
}
