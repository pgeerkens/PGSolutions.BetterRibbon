////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using IStrings = IControlStrings;

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IGroupModel)]
    public interface IGroupModel {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IStrings Strings {
            [Description("Gets the IControlStrings for this control.")]
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

        /// <summary>Gets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Gets whether or not inactive controls should be visible on the Ribbon.")]
        bool ShowInactive { get; }

        /// <summary>Sets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Sets whether or not inactive controls should be visible on the Ribbon.")]
        void SetShowInactive(bool showInactive);

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IGroupModel Attach(string controlId);

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();

        /// <summary>Detaches this Ribbon Group, and all child models, from their view-models.</summary>
        [Description("Detaches this Ribbon Group, and all child models, from their view-models.")]
        void Detach();
    }
}
