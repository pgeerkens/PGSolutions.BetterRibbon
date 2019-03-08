////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>This interface supplies new ribbon-model obejcts, unattached to any view-model object.</summary>
    /// <remarks>
    /// The model objects returned by this interface will not respond to any Ribbon callbacks until 
    /// they have been attached to a view-model object. The 'stringsId' supplied to the constructor
    /// is used to retrieve display strings from the client-supplied IResourceLoader - and need not
    /// be identical to the controlId subsequently used to identify the attahced view-model control.
    /// 
    /// The <see cref="SuppressMessage"/> attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file, as commentary on a practice often seen as a C#
    /// anti-pattern. Although non-standard C# practice, these "optional parameters with default 
    /// values" usages are (believed to be) the only means of implementing functionality equivalent
    /// to "overrides" in a COM-compatible way.
    /// </remarks>
    [Description(
@"This interface supplies new ribbon-model obejcts, unattached to any view-model object.

The model objects returned by this interface will not respond to any Ribbon callbacks until 
they have been attached to a view-model object. The 'stringsId' supplied to the constructor
is used to retrieve display strings from the client-supplied IResourceLoader - and need not
relate to the controlId subsequently used to identify the attahced view-model control.

In the current implementation interfaces IModelFactory and IModelServer are provided by the
same underlying object. It is the intent that they should always work properly in conjunction
with each other."
    )]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IModelFactory)]
    public interface IModelFactory {
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
        [DispId(4), Description("Returns a new ImageObject from the supplied IPictureDisp.")]
        IImageObject NewImageObject(IPictureDisp image);

        /// <summary>Returns a new <see cref="IImageObject"/> from the supplied MSO image name.</summary>
        [DispId(5), Description("Returns a new ImageObject from the supplied MSO image name.")]
        IImageObject NewImageObjectMso(string imageMso);

        /// <summary>Returns a new, unattached, ribbon Group model.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(6), Description("Returns a new, unattached, ribbon Group model.")]
        IGroupModel NewGroupModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon Button model.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(7), Description("Returns a new, unattached, ribbon Button model.")]
        IButtonModel NewButtonModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon Toggle model. Attachable to either a ToggleButton or a CheckBox</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(8), Description("Returns a new, unattached, ribbon Toggle model. Attachable to either a ToggleButton or a CheckBox")]
        IToggleModel NewToggleModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon EditBOx model.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(9), Description("Returns a new, unattached, ribbon EditBox model.")]
        IEditBoxModel NewEditBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon DropDown model that supports a dynamic selection list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(10), Description("Returns a new, unattached, ribbon DropDown model that supports a dynamic selection list.")]
        IDropDownModel NewDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon DropDown model that supports a static selection list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(20), Description("Returns a new, unattached, ribbon DropDown model that supports a static selection list.")]
        IStaticDropDownModel NewStaticDropDownModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon ComboBox model that supports a dynamic selection list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(11), Description("Returns a new, unattached, ribbon ComboBox model that supports a dynamic selection list.")]
        IComboBoxModel NewComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon ComboBox model that supports a static selection list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(21), Description("Returns a new, unattached, ribbon ComboBox model that supports a static selection list.")]
        IStaticComboBoxModel NewStaticComboBoxModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon LabelControl model.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(12), Description("Returns a new, unattached, ribbon LabelControl model.")]
        ILabelControlModel NewLabelControlModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon menu model.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(13), Description("Returns a new, unattached, ribbon Menu model.")]
        IMenuModel NewMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Toggle)Button control.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(14), Description("Returns a new model for a Split(Toggle)Button control.")]
        ISplitToggleButtonModel NewSplitToggleButtonModel(string splitStringsId, string menuStringsId,
                string toggleStringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for a Split(Press)Button control.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(15), Description("Returns a new model for a Split(Press)Button control.")]
        ISplitPressButtonModel NewSplitPressButtonModel(string splitStringsId, string menuStringsId,
                string buttonStringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new model for an Item suitable for use in a DropDown, ComboBox or X selection list.</summary>
        /// <param name="controlId">The ID of the new {ISelectableItem} to be returned.</param>
        [DispId(19), Description("Returns a new model for an Item suitable for use in a DropDown, ComboBox or X selection list.")]
        ISelectableItemModel NewSelectableModel(string controlID);

        /// <summary>Returns a new, unattached, ribbon Gallery model that supports a dynamic selection list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(22), Description("Returns a new, unattached, ribbon Gallery model that supports a dynamic selection list.")]
        IGalleryModel NewGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon Gallery model that supports a static selection list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(23), Description("Returns a new, unattached, ribbon Gallery model that supports a static selection list.")]
        IStaticGalleryModel NewStaticGalleryModel(string stringsId, bool isEnabled = true, bool isVisible = true);

        /// <summary>Returns a new, unattached, ribbon MenuSeparator model.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(24), Description("Returns a new, unattached, ribbon MenuSeparator.")]
        IMenuSeparatorModel NewMenuSeparatorModel(string stringsId, bool isEnabled = true, bool isVisible = true);

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

        /// <summary>Returns a new, unattached, ribbon DynamicMenu model.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(29), Description("Returns a new ribbon DynamicMenu model.")]
        IDynamicMenuModel NewDynamicMenuModel(string stringsId, bool isEnabled = true, bool isVisible = true);
    }
}
