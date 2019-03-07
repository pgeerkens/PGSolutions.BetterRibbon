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
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>This alternative interface returns model controls, automatically attahced to the eponymous view-model control if it exists.</summary>
    /// <remarks>
    /// The {SuppressMessage} attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file, as commentary on a practice often seen as a C#
    /// anti-pattern. Although non-standard C# practice, these "optional parameters with default 
    /// values" usages are (believed to be) the only means of implementing functionality equivalent
    /// to "overrides" in a COM-compatible way.
    /// </remarks>
    [Description(
@"This alternative interface returns model controls, automatically attached to the eponymous view-
model control if it exists.

To effectively use this (alternative) interface the stringsId for each model control used in
retrieving its display strings must be the same as the controlId for the view-model being
attached to - as only one string is provided. This does not prevent mised usage

In the current implementation interfaces IModelFactory and IModelServer are provided by the
same underlying object. It is the intent that they should always work properly in conjunction
with each other."
    )]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IModelServer)]
    public interface IModelServer {
        ///// <summary>Queues a refresh of the PGSolutions Ribbon Tab.</summary>
        //[DispId(1), Description("Queues a refresh of the PGSolutions Ribbon Tab.")]
        //void Invalidate();

        /// <summary>Returns a new <see cref="IStrings"/> constructed from the supplied strings.</summary>
        [DispId(2), Description("Returns a new IControlStrings constructed from the supplied strings.")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IStrings NewControlStrings(string label, string screenTip="", string superTip="", string keyTip="");

        /// <summary>Returns a new <see cref="IStrings2"/> constructed from the supplied strings.</summary>
        [DispId(18), Description("Returns a new IControlStrings2 constructed from the supplied strings.")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IStrings2 NewControlStrings2(string label, string screenTip="", string superTip="", string keyTip="",
                string description="");

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

        /// <summary>Returns a new ribbon Group model, attached to the named view-model ribbon group.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(6), Description("Returns a new ribbon Group model, attached to the named view-model ribbon group.")]
        IGroupModel GetGroupModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon Button model, attached to the named view-model ribbon button.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(7), Description("Returns a new ribbon Button model, attached to the named view-model ribbon button.")]
        IButtonModel GetButtonModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon ToggleButton model, attached to the named view-model ribbon toggleButton.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(8), Description("Returns a new ribbon ToggleButton model, attached to the named view-model ribbon toggleButton.")]
        IToggleModel GetToggleModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon EditBOx model, attached to the named view-model ribbon editBox.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(9), Description("Returns a new ribbon EditBOx model, attached to the named view-model ribbon editBox.")]
        IEditBoxModel GetEditBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon DropDown model, attached to the named view-model ribbon dropDOwn.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(10), Description("Returns a new ribbon DropDown model, attached to the named view-model ribbon dropDOwn.")]
        IDropDownModel GetDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon DropDown model, attached to the named view-model static ribbon dropDown.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(20), Description("Returns a new ribbon DropDown model, attached to the named view-model static ribbon dropDown.")]
        IStaticDropDownModel GetStaticDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon ComboBox model, attached to the named view-model ribbon comboBox.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(11), Description("Returns a new ribbon ComboBox model, attached to the named view-model static ribbon comboBox.")]
        IComboBoxModel GetComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon ComboBox model, attached to the named view-model static ribbon comboBox.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(21), Description("Returns a new ribbon ComboBox model, attached to the named view-model static ribbon comboBox.")]
        IStaticComboBoxModel GetStaticComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon Label model, attached to the named view-model static ribbon labelControl.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(12), Description("Returns a new ribbon Label model, attached to the named view-model static ribbon labelControl.")]
        ILabelControlModel GetLabelControlModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon Menu model, attached to the named view-model static ribbon menu.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(13), Description("Returns a new ribbon Menu model, attached to the named view-model static ribbon menu.")]
        IMenuModel GetMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Toggle)Button control.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "string")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(14), Description("Returns a new model for a Split(Toggle)Button control.")]
        ISplitToggleButtonModel GetSplitToggleButtonModel(string splitStringId, string menuStringId,
                string toggleStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Press)Button control.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "string")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(15), Description("Returns a new model for a Split(Press)Button control.")]
        ISplitPressButtonModel GetSplitPressButtonModel(string splitStringId, string menuStringId,
                string buttonStringId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon SelectableItem model, attached to the named view-model ribbon item.</summary>
        /// <param name="controlId">The ID of the new {ISelectableItem} to be returned.</param>
        [DispId(19), Description("Returns a new ribbon SelectableItem model, attached to the named view-model static ribbon item.")]
        ISelectableItemModel GetSelectableModel(string controlID);

        /// <summary>Returns a new ribbon Gallery model, attached to the named view-model ribbon gallery.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(22), Description("Returns a new ribbon Gallery model, attached to the named view-model static ribbon gallery.")]
        IGalleryModel GetGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon Gallery model, attached to the named view-model static ribbon gallery.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(23), Description("Returns a new ribbon Gallery model, attached to the named view-model static ribbon gallery.")]
        IStaticGalleryModel GetStaticGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new ribbon MenuSeparator model, attached to the named view-model ribbon menuSeparator.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(24), Description("Returns a new ribbon MenuSeparator model, attached to the named view-model ribbon menuSeparator.")]
        IMenuSeparatorModel GetMenuSeparatorModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns an <see cref="IStrings"/> requested from the client-supplied <see cref="IResourceLoader"/> for the supplied <paramref name="id"/>.</summary>
        [DispId(25), Description("Returns an IControlStrings requested from the client-supplied IResourceLoader for the supplied id..")]
        IStrings GetStrings(string id);

        /// <summary>Returns an <see cref="IStrings2"/> requested from the client-supplied <see cref="IResourceLoader"/> for the supplied <paramref name="id"/>.</summary>
        [DispId(26), Description("Returns an IControlStrings2 requested from the client-supplied IResourceLoader for the supplied id.")]
        IStrings2 GetStrings2(string id);

        /// <summary>Returns the supplied <see cref="IPictureDisp"/> wrapped as an <see cref="IImageObject"/>.</summary>
        [DispId(27), Description("Returns the supplied IPictureDisp wrapped as an IImageObject.")]
        IImageObject GetImage(IPictureDisp image);

        /// <summary>Returns the supplied MSO name-string wrapped as an <see cref="IImageObject"/>.</summary>
        [DispId(28), Description("Returns the supplied MSO name-string wrapped as an IImageObject.")]
        IImageObject GetImageMso(string imageMso);

        /// <summary>Returns a new ribbon DynamicMenu model, attached to the named view-model ribbon dynamicMenu.</summary>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "strings")]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(29), Description("Returns a new ribbon DynamicMenu model, attached to the named view-model ribbon dynamicMenu.")]
        IDynamicMenuModel GetDynamicMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);
    }
}
