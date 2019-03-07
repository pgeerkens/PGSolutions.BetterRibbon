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
        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        [DispId(4)]
        string Label {
            [Description("Gets the IControlStrings for this control.")]
            get; set;
        }
        /// <summary>Gets the ScreenTip (concise hover-help) for this control.</summary>
        [DispId(17)]
        string ScreenTip {
            [Description("Gets the ScreenTip (concise hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the SuperTip (expanded hover-help) for this control.</summary>
        [DispId(18)]
        string SuperTip {
            [Description("Gets the SuperTip (expanded hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the KeyTip (keyboard shortcut) for this control.</summary>
        [DispId(19)]
        string KeyTip {
            [Description("Gets the KeyTip (keyboard shortcut) for this control.")]
            get; set;
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

        /// <summary>Gets or sets Whether this control displays an image.</summary>
        [DispId(10)]
        bool   ShowImage  {
            [Description("Gets or sets Whether this control displays an image.")]
            get; set; }
    }

    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISplitToggleButtonModel)]
    public interface ISplitToggleButtonModel: ISplitButtonModel {
        /// <summary>True exaactly when this control is in the Pressed state. Default value.</summary>
        [DispId(12)]    
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
        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        [DispId(4)]
        new string Label {
            [Description("Gets the IControlStrings for this control.")]
            get; set;
        }
        /// <summary>Gets the ScreenTip (concise hover-help) for this control.</summary>
        [DispId(17)]
        new string ScreenTip {
            [Description("Gets the ScreenTip (concise hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the SuperTip (expanded hover-help) for this control.</summary>
        [DispId(18)]
        new string SuperTip {
            [Description("Gets the SuperTip (expanded hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the KeyTip (keyboard shortcut) for this control.</summary>
        [DispId(19)]
        new string KeyTip {
            [Description("Gets the KeyTip (keyboard shortcut) for this control.")]
            get; set;
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

        /// <summary>Gets or sets Whether this control displays an image.</summary>
        [DispId(10)]
        new bool   ShowImage  {
            [Description("Gets or sets Whether this control displays an image.")]
            get; set; }
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
        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        [DispId(4)]
        new string Label {
            [Description("Gets the IControlStrings for this control.")]
            get; set;
        }
        /// <summary>Gets the ScreenTip (concise hover-help) for this control.</summary>
        [DispId(17)]
        new string ScreenTip {
            [Description("Gets the ScreenTip (concise hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the SuperTip (expanded hover-help) for this control.</summary>
        [DispId(18)]
        new string SuperTip {
            [Description("Gets the SuperTip (expanded hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the KeyTip (keyboard shortcut) for this control.</summary>
        [DispId(19)]
        new string KeyTip {
            [Description("Gets the KeyTip (keyboard shortcut) for this control.")]
            get; set;
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

        /// <summary>Gets or sets Whether this control displays an image.</summary>
        [DispId(10)]
        new bool   ShowImage  {
            [Description("Gets or sets Whether this control displays an image.")]
            get; set; }
    }
}
