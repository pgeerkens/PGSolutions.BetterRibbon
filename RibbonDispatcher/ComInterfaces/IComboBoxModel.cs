////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IComboBoxModel)]
    public interface IComboBoxModel {
        /// <summary>Gets or sets the content of the <see cref="EditBoxVM"/>. Default value.</summary>
        [DispId(0)]
        string Text {
            [Description("Gets or sets the content of the EditBOxVM. Default value.")]
            get; set;
        }

        #region IActivable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [DispId(1),Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IComboBoxModel Attach(string controlId);

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

        /// <summary>Adds the specified <see cref="ISelectableItem"/> to the available options in the drop-down list.</summary>
        [DispId( 8),Description("Adds the specified ISelectableItem to the available options in the drop-down list.")]
        IComboBoxModel ClearList();

        /// <summary>Adds the specified <see cref="ISelectableItem"/> to the available options in the drop-down list.</summary>
        [DispId( 9),Description("Adds the specified ISelectableItem to the available options in the drop-down list.")]
        IComboBoxModel AddSelectableModel(IStaticItemVM selectableModel);
    }
}
