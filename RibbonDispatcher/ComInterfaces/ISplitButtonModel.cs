////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [CLSCompliant(true)]
    //[ComVisible(true)]
    //[InterfaceType(ComInterfaceType.InterfaceIsDual)]
    //[Guid(Guids.ISplitButtonModel)]
    public interface ISplitButtonModel {
        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        IControlStrings Strings {
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
        /// <summary>.</summary>
        bool IsLarge {
            [Description(".")]
            get; set;
        }

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
    }

    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISplitToggleButtonModel)]
    public interface ISplitToggleButtonModel: ISplitButtonModel {
        /// <summary>Gets or sets whether the control is pressed.</summary>
        bool IsPressed {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        ISplitToggleButtonModel Attach(string controlId);

        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        new IControlStrings Strings {
            [Description("Gets the IControlStrings for this control.")]
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
        /// <summary>.</summary>
        new bool IsLarge {
            [Description(".")]
            get; set;
        }

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        new void Invalidate();
    }

    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISplitPressButtonModel)]
    public interface ISplitPressButtonModel: ISplitButtonModel {
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        ISplitPressButtonModel Attach(string controlId);

        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        new IControlStrings Strings {
            [Description("Gets the IControlStrings for this control.")]
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
        /// <summary>.</summary>
        new bool IsLarge {
            [Description(".")]
            get; set;
        }

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        new void Invalidate();
    }
}
