////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using IStrings  = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>The main interface for VBA to access the Ribbon dispatcher.</summary>
    [Description("The main interface for VBA to access the Ribbon dispatcher.")]
        [CLSCompliant(true)][ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IModelFactory)]
    public interface IModelFactory {
        ///// <summary>Queues a refresh of the PGSolutions Ribbon Tab.</summary>
        //[DispId(1), Description("Queues a refresh of the PGSolutions Ribbon Tab.")]
        //void Invalidate();

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(2), Description(".")]
        IStrings NewControlStrings(string label, string screenTip, string superTip, string keyTip);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(18), Description(".")]
        IStrings2 NewControlStrings2(string label, string screenTip, string superTip, string keyTip,
                string description);

        /// <summary>Deactivate the specified control, detaching any attached data source.</summary>
        /// <param name="controlId">The ID of the control to be detached.</param>
        [DispId(3), Description("Deactivate the specified control, detaching any attached data source.")]
        void DetachProxy(string controlId);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(4), Description(".")]
        IGroupModel NewGroupModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(5), Description(".")]
        IButtonModel NewButtonModel(string stringsId,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(6), Description(".")]
        IButtonModel NewButtonModelMso(string stringsId,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(7), Description(".")]
        IToggleModel NewToggleModel(string stringsId,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(8), Description(".")]
        IToggleModel NewToggleModelMso(string stringsId,
                string imageMso = "MacroSecurity", bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(9), Description(".")]
        IEditBoxModel NewEditBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(10), Description(".")]
        IDropDownModel NewDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(20), Description(".")]
        IStaticDropDownModel NewStaticDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(11), Description(".")]
        IComboBoxModel NewComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(21), Description(".")]
        IStaticComboBoxModel NewStaticComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(22), Description(".")]
        IGalleryModel NewGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(23), Description(".")]
        IStaticGalleryModel NewStaticGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(12), Description(".")]
        ILabelModel NewLabelModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(13), Description(".")]
        IMenuModel NewMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Toggle)Button control.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "string")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(14), Description("Returns a new model for a Split(Toggle)Button control.")]
        ISplitButtonModel NewSplitToggleButtonModel(string splitStringId, string menuStringId,
                string toggleStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Press)Button control.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "string")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(15), Description("Returns a new model for a Split(Press)Button control.")]
        ISplitButtonModel NewSplitPressButtonModel(string splitStringId, string menuStringId,
                string buttonStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        /// <param name="controlId">The ID of the new {ISelectableItem} to be returned.</param>
        [DispId(19), Description(".")]
        ISelectableItemModel NewSelectableModel(string controlID);
    }
}
