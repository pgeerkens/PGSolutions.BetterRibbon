////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

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
        IRibbonControlStrings GetStrings(string controlId);

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        T Add<T, TSource>(T ctrl) where T : RibbonCommon<TSource> where TSource : class, IRibbonCommonSource;

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonGroupViewModel NewRibbonGroup(string controlId);

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonButton NewRibbonButton(string controlId);

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonToggleButton NewRibbonToggle(string controlId);

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonCheckBox NewRibbonCheckBox(string controlId);

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonDropDown NewRibbonDropDown(string controlId);

        /// <summary>Returns a new {SelectableItem} from a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        SelectableItem NewSelectableItem(string controlId);

        /// <summary>Returns a new {ResourceLoader} object.</summary>
        IResourceLoader NewResourceLoader();

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IRibbonControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "");
    }
}
