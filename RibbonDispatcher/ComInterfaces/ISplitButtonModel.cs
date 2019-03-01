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
    public interface ISplitButtonModel {
        #region IActivable implementation
        /// <summary>.</summary>
        [DispId(2),Description(".")]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [DispId(3),Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
        #endregion

        #region IControl implementation
        /// <summary>Gets the {IControlStrings} for this control.</summary>
        [DispId(4)]
        IControlStrings Strings {
            [Description("Gets the {IControlStrings} for this control.")]
            get;
        }
        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(5)]
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(6)]
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }
        #endregion

        #region ISizeable implementation
        /// <summary>.</summary>
        [DispId(7)]
        bool   IsLarge    {
            [Description(".")]
            get; set; }
        #endregion
    }

    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISplitToggleButtonModel)]
    public interface ISplitToggleButtonModel: ISplitButtonModel {
        /// <summary>True exaactly when this control is in the Pressed state. Default value.</summary>
        [DispId(0)]    
        bool IsPressed {
            [Description("True exaactly when this control is in the Pressed state. Default value.")]
            get; set;
        }

        #region IActivable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [DispId(1),Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        ISplitToggleButtonModel Attach(string controlId);

        /// <summary>.</summary>
        [DispId(2),Description(".")]
        new void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [DispId(3),Description("Queues a request for this control to be refreshed.")]
        new void Invalidate();
        #endregion

        #region IControl implementation
        /// <summary>Gets the {IControlStrings} for this control.</summary>
        [DispId(4)]
        new IControlStrings Strings {
            [Description("Gets the {IControlStrings} for this control.")]
            get;
        }
        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(5)]
        new bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(6)]
        new bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }
        #endregion

        #region ISizeable implementation
        /// <summary>.</summary>
        [DispId(7)]
        new bool   IsLarge    {
            [Description(".")]
            get; set; }
        #endregion

        [DispId(8)]
        IMenuModel MenuModel { get; }

        [DispId(9)]
        IToggleModel ToggleModel { get; }
    }

    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISplitPressButtonModel)]
    public interface ISplitPressButtonModel: ISplitButtonModel {
        #region IActivable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [DispId(1),Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        ISplitPressButtonModel Attach(string controlId);

        /// <summary>.</summary>
        [DispId(2),Description(".")]
        new void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [DispId(3),Description("Queues a request for this control to be refreshed.")]
        new void Invalidate();
        #endregion

        #region IControl implementation
        /// <summary>Gets the {IControlStrings} for this control.</summary>
        [DispId(4)]
        new IControlStrings Strings {
            [Description("Gets the {IControlStrings} for this control.")]
            get;
        }
        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(5)]
        new bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(6)]
        new bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }
        #endregion

        #region ISizeable implementation
        /// <summary>.</summary>
        [DispId(7)]
        new bool   IsLarge    {
            [Description(".")]
            get; set; }
        #endregion

        [DispId(8)]
        IMenuModel MenuModel { get; }

        [DispId(9)]
        IButtonModel ButtonModel { get; }
    }
}
