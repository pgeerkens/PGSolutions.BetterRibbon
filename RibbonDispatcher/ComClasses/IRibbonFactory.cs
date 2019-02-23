////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The factory interface for the Ribbon Dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonFactory)]
    public interface IRibbonFactory {
        /// <summary>TODO</summary>
        IResourceManager ResourceManager { get; }

        /// <summary>.</summary>
        /// <param name="controlId"></param>
        [Description("")]
        IControlStrings GetStrings(string controlId);

        //[SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        //T Add<T, TSource>(T ctrl) where T : AbstractControlVM<TSource> where TSource : class, IRibbonCommonSource;

        ///// <summary>TODO</summary>
        //TControl GetControl<TControl>(string controlId) where TControl : class, IRibbonControlVM;

        ///// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        //GroupVM NewGroup(string controlId);

        ///// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses a custom Image (or none).</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        //ButtonVM NewButton(string controlId);

        ///// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses a custom Image (or none).</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        //ToggleButtonVM NewToggleButton(string controlId);

        ///// <summary>Returns a new Ribbon CheckBoxVM ViewModel instance.</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        //CheckBoxVM NewCheckBox(string controlId);

        ///// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        //DropDownVM NewDropDown(string controlId);

        ///// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        //EditBoxVM NewEditBox(string controlId);

        ///// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        //ComboBoxVM NewComboBox(string controlId);

        ///// <summary>Returns a new {SelectableItem} from a custom Image (or none).</summary>
        //[SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        //SelectableItem NewSelectableItem(string controlId);

        /// <summary>Returns a new {ResourceLoader} object.</summary>
        IResourceLoader NewResourceLoader();

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "");

        object LoadImage(string imageId);
    }
}
