using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.AbstractCOM;
using PGSolutions.RibbonDispatcher.Concrete;

namespace PGSolutions.RibbonDispatcher {
    using static RdControlSize;

    /// <summary>The factory interface for the Ribbon Dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonFactory)]
    public interface IRibbonFactory {
        IResourceManager ResourceManager { get; }

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [DispId(DispIds.NewRibbonGroup)]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonGroup NewRibbonGroup(string ItemId, bool Visible = true, bool Enabled = true);

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonButton)]
        RibbonButton NewRibbonButton(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize   Size            = rdLarge,
            IPictureDisp    Image           = null,
            bool            ShowImage       = false,
            bool            ShowLabel       = true
        );
        /// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses an Office built-in Image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonButtonMso)]
        RibbonButton NewRibbonButtonMso(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize   Size            = rdLarge,
            string          ImageMso        = "MacroSecurity",  // This one get's people's attention ;-)
            bool            ShowImage       = false,
            bool            ShowLabel       = true
        );

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonToggle)]
        RibbonToggleButton NewRibbonToggle(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize   Size            = rdLarge,
            IPictureDisp    Image           = null,
            bool            ShowImage       = false,
            bool            ShowLabel       = true
        );
        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses an Office built-in Image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonToggleMso)]
        RibbonToggleButton NewRibbonToggleMso(string ItemId, bool Visible = true, bool Enabled = true,
            RdControlSize   Size            = rdLarge,
            string          ImageMso        = "MacroSecurity",  // This one gets people's attention ;-)
            bool            ShowImage       = false,
            bool            ShowLabel       = true
        );

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonCheckBox)]
        RibbonCheckBox NewRibbonCheckBox(string ItemId, bool Visible = true, bool Enabled = true);

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonDropDown)]
        RibbonDropDown NewRibbonDropDown(string ItemId, bool Visible = true, bool Enabled = true);

        /// <summary>Returns a new {SelectableItem} from a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(DispIds.NewSelectableItem)]
        SelectableItem NewSelectableItem(string ItemId, IPictureDisp Image = null);

        /// <summary>Returns a new {SelectableItem} from an Office built-in Image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(DispIds.NewSelectableItemMso)]
        SelectableItem NewSelectableItemMso(string ItemId, string ImageMso = "MacroSecurity");

        /// <summary>Returns a new {ResourceLoader} object.</summary>
        [DispId(DispIds.NewResourceLoader)]
        IResourceLoader NewResourceLoader();
    }
}
