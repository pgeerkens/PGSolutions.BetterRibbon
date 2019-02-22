////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The main interface for VBA to access the Ribbon dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonDispatcher)]
    public interface IRibbonDispatcher {
        /// <summary>Queues a refresh of the PGSolutions Ribbon Tab.</summary>
        [Description("Queues a refresh of the PGSolutions Ribbon Tab.")]
        //[DispId(2)]
        void Invalidate();

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        //[DispId(3)]
        IRibbonControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "");

        /// <summary>Deactivate the specified control, detaching any attached data source.</summary>
        /// <param name="controlId">The ID of the control to be detached.</param>
        [Description("Deactivate the specified control, detaching any attached data source.")]
        //[DispId(4)]
        void DetachProxy(string controlId);

        /// <summary>.</summary>
        /// <param name="controlId">The ID of the new {ISelectableItem} to be returned.</param>
        [Description(".")]
        ISelectableItemModel NewSelectableModel(string controlID, IRibbonControlStrings strings);

        /// <summary>.</summary>
        [Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        IRibbonButtonModel NewButtonModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IRibbonButtonModel NewButtonModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IRibbonToggleModel NewToggleModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IRibbonToggleModel NewToggleModelMso(IRibbonControlStrings strings,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IRibbonDropDownModel NewDropDownModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IEditBoxModel NewEditBoxModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IRibbonGroupModel NewGroupModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true);
    }
}
