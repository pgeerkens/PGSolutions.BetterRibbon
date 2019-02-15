////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

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

        T Add<T, TSource>(T ctrl) where T : RibbonCommon<TSource> where TSource : class, IRibbonCommonSource;

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [DispId(DispIds.NewRibbonGroup)]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonGroupViewModel NewRibbonGroup(string itemId, bool visible = true, bool enabled = true);

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonButton)]
        RibbonButton NewRibbonButton(string itemId, bool visible = true, bool enabled = true,
            bool         isLarge   = true,
            IPictureDisp image     = null,
            bool         showImage = true,
            bool         showLabel = true
        );
        /// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses an Office built-in Image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonButtonMso)]
        RibbonButton NewRibbonButtonMso(string itemId, bool visible = true, bool enabled = true,
            bool         isLarge   = true,
            string       imageMso  = "MacroSecurity",  // This one gets people's attention ;-)
            bool         showImage = true,
            bool         showLabel = true
        );

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonToggle)]
        RibbonToggleButton NewRibbonToggle(string itemId, bool visible = true, bool enabled = true,
            bool         isLarge   = true,
            IPictureDisp image     = null,
            bool         showImage = true,
            bool         showLabel = true
        );
        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses an Office built-in Image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonToggleMso)]
        RibbonToggleButton NewRibbonToggleMso(string itemId, bool visible = true, bool enabled = true,
            bool         isLarge   = true,
            string       imageMso  = "MacroSecurity",  // This one gets people's attention ;-)
            bool         showImage = true,
            bool         showLabel = true
        );

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonCheckBox)]
        RibbonCheckBox NewRibbonCheckBox(string itemId, bool visible = true, bool enabled = true);

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(DispIds.NewRibbonDropDown)]
        RibbonDropDown NewRibbonDropDown(string itemId, bool visible = true, bool enabled = true);

        /// <summary>Returns a new {SelectableItem} from a custom Image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(DispIds.NewSelectableItem)]
        SelectableItem NewSelectableItem(string itemId, IPictureDisp image = null);

        /// <summary>Returns a new {SelectableItem} from an Office built-in Image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(DispIds.NewSelectableItemMso)]
        SelectableItem NewSelectableItemMso(string itemId, string imageMso = "MacroSecurity");

        /// <summary>Returns a new {ResourceLoader} object.</summary>
        [DispId(DispIds.NewResourceLoader)]
        IResourceLoader NewResourceLoader();

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(DispIds.NewControlStrings)]
        IRibbonControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "");
    }
}
namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    internal static partial class DispIds {
        public const int NewRibbonGroup         = 1;
        public const int NewRibbonButton        = 1 + NewRibbonGroup;
        public const int NewRibbonButtonMso     = 1 + NewRibbonButton;
        public const int NewRibbonToggle        = 1 + NewRibbonButtonMso;
        public const int NewRibbonToggleMso     = 1 + NewRibbonToggle;
        public const int NewRibbonCheckBox      = 1 + NewRibbonToggleMso;
        public const int NewRibbonDropDown      = 1 + NewRibbonCheckBox;
        public const int NewSelectableItem      = 1 + NewRibbonDropDown;
        public const int NewSelectableItemMso   = 1 + NewSelectableItem;
        public const int NewResourceLoader      = 1 + NewSelectableItemMso;
        public const int NewControlStrings      = 1 + NewResourceLoader;
    }
}
