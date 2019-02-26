////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComClasses;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using IStrings  = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>The main interface for VBA to access the Ribbon dispatcher.</summary>
    [CLSCompliant(true)][ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IModelFactory)]
    public interface IModelFactory {
        ///// <summary>Queues a refresh of the PGSolutions Ribbon Tab.</summary>
        //[DispId(1), Description("Queues a refresh of the PGSolutions Ribbon Tab.")]
        //void Invalidate();

        /// <summary>.</summary>
        [DispId(2), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IStrings NewControlStrings(string label, string screenTip, string superTip,
                string keyTip);

        /// <summary>.</summary>
        [DispId(18), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IStrings2 NewControlStrings2(string label, string screenTip, string superTip,
                string keyTip, string description);

        ///// <summary>Deactivate the specified control, detaching any attached data source.</summary>
        ///// <param name="controlId">The ID of the control to be detached.</param>
        //[DispId(3), Description("Deactivate the specified control, detaching any attached data source.")]
        //void DetachProxy(string controlId);

        /// <summary>.</summary>
        [DispId(4), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IGroupModel NewGroupModel(IStrings strings,
                bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(5), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        IButtonModel NewButtonModel(string stringsId,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(6), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IButtonModel NewButtonModelMso(string stringsId,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(7), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IToggleModel NewToggleModel(string stringsId,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(8), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IToggleModel NewToggleModelMso(string stringsId,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(9), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IEditBoxModel NewEditBoxModel(string stringsId,
                bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(10), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IDropDownModel NewDropDownModel(string stringsId,
                bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(11), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IComboBoxModel NewComboBoxModel(string stringsId,
                bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(12), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        ILabelModel NewLabelModel(string stringsId,
                bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(13), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IMenuModel NewMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(14), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        ISplitButtonModel NewSplitToggleButtonModel(string splitStringId, string menuStringId,
                string toggleStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(15), Description(".")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        ISplitButtonModel NewSplitPressButtonModel(string splitStringId, string menuStringId,
                string buttonStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        /// <param name="controlId">The ID of the new {ISelectableItem} to be returned.</param>
        [DispId(19), Description(".")]
        ISelectableItemModel NewSelectableModel(string controlID, IStrings strings);
    }

    public interface IModelFactoryInternal: IModelFactory {
        /// <summary>.</summary>
        ViewModelFactory ViewModelFactory { get; }

        /// <summary>.</summary>
        IResourceManager ResourceManager { get; }
    }
}
