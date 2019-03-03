////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ViewModels;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using IStrings  = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>The main interface for VBA to access the Ribbon dispatcher.</summary>
    /// <remarks>
    /// The {SuppressMessage} attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file, as commentary on a practice often seen as a C#
    /// anti-pattern. Although non-standard C# practice, these "optional parameters with default 
    /// values" usages are (believed to be) the only means of implementing functionality equivalent
    /// to "overrides" in a COM-compatible way.
    /// </remarks>
    [Description("The main interface for VBA to access the Ribbon dispatcher.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IModelFactory)]
    public interface IModelFactory {
        ///// <summary>Queues a refresh of the PGSolutions Ribbon Tab.</summary>
        //[DispId(1), Description("Queues a refresh of the PGSolutions Ribbon Tab.")]
        //void Invalidate();

        /// <summary>.</summary>
        [DispId(2), Description(".")]
        IStrings NewControlStrings(string label, string screenTip, string superTip, string keyTip);

        /// <summary>.</summary>
        [DispId(18), Description(".")]
        IStrings2 NewControlStrings2(string label, string screenTip, string superTip, string keyTip,
                string description);

        /// <summary>Deactivate the specified control, detaching any attached data source.</summary>
        /// <param name="controlId">The ID of the control to be detached.</param>
        [DispId(3), Description("Deactivate the specified control, detaching any attached data source.")]
        void DetachProxy(string controlId);

        /// <summary>Returns a new <see cref="IImageObject"/> from the supplied <see cref="IPictureDisp"/>.</summary>
        [SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings")]
        [DispId(4), Description("Returns a new ImageObject from the supplied IPictureDisp.")]
        IImageObject NewImageObject(IPictureDisp image);

        /// <summary>Returns a new <see cref="IImageObject"/> from the supplied MSO image name.</summary>
        [SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings")]
        [DispId(5), Description("Returns a new ImageObject from the supplied MSO image name.")]
        IImageObject NewImageObjectMso(string imageMso);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(6), Description(".")]
        IGroupModel NewGroupModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(7), Description(".")]
        IButtonModel NewButtonModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(8), Description(".")]
        IToggleModel NewToggleModel(string stringsId, bool isEnabled = true, bool isVisible = true);

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
        [DispId(12), Description(".")]
        ILabelControlModel NewLabelControlModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(13), Description(".")]
        IMenuModel NewMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Toggle)Button control.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "string")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(14), Description("Returns a new model for a Split(Toggle)Button control.")]
        ISplitToggleButtonModel NewSplitToggleButtonModel(string splitStringId, string menuStringId,
                string toggleStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Press)Button control.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "string")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(15), Description("Returns a new model for a Split(Press)Button control.")]
        ISplitPressButtonModel NewSplitPressButtonModel(string splitStringId, string menuStringId,
                string buttonStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        /// <param name="controlId">The ID of the new {ISelectableItem} to be returned.</param>
        [DispId(19), Description(".")]
        ISelectableItemModel NewSelectableModel(string controlID);

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
        [DispId(24), Description(".")]
        IMenuSeparatorModel NewMenuSeparatorModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>.</summary>
        [DispId(25), Description(".")]
        IStrings GetStrings(string id);

        /// <summary>.</summary>
        [DispId(26), Description(".")]
        IStrings2 GetStrings2(string id);

        /// <summary>.</summary>
        [DispId(27), Description(".")]
        IImageObject GetImage(IPictureDisp image);

        /// <summary>.</summary>
        [DispId(28), Description(".")]
        IImageObject GetImage(string imageMso);

        /// <summary>.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(29), Description(".")]
        IDynamicMenuModel NewDynamicMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);
    }
}
