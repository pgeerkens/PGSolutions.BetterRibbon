using System;
using System.Globalization;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ControlMixins;
using PGSolutions.RibbonDispatcher.AbstractCOM;

namespace PGSolutions.RibbonDispatcher.Concrete {

    /// <summary>Implementation of (all) the callbacks for the Fluent Ribbon; for .NET clients.</summary>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [ComDefaultInterface(typeof(IRibbonViewModel))]
    [Guid(Guids.AbstractDispatcher)]
    public abstract class AbstractRibbonViewModel : IRibbonViewModel {
        /// <summary>TODO</summary>
        protected void InitializeRibbonFactory(IRibbonUI RibbonUI, IResourceManager ResourceManager) 
            => _ribbonFactory = new RibbonFactory(RibbonUI, ResourceManager);

        /// <summary>TODO</summary>
        public object LoadImage(string imageId) => _ribbonFactory.LoadImage(imageId);

        /// <inheritdoc/>
        public IRibbonFactory RibbonFactory => _ribbonFactory; private RibbonFactory _ribbonFactory;

        /// <summary>TODO</summary>
        public void Invalidate()                            => _ribbonFactory.Invalidate();
        /// <summary>TODO</summary>
        public void InvalidateControl(string ControlId)     => _ribbonFactory.InvalidateControl(ControlId);
        /// <summary>TODO</summary>
        public void InvalidateControlMso(string ControlId)  => _ribbonFactory.InvalidateControlMso(ControlId);
        /// <summary>TODO</summary>
        public void ActivateTab(string ControlId)           => _ribbonFactory.ActivateTab(ControlId);

        /// <summary>All of the defined controls.</summary>
        private IRibbonCommon    Controls    (string controlId) => _ribbonFactory.Controls.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {ISizeableMixin} interface.</summary>
        private ISizeableMixin   Sizeables   (string controlId) => _ribbonFactory.Sizeables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {IActionableMixin} interface.</summary>
        private IClickableMixin  Actionables (string controlId) => _ribbonFactory.Actionables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {IToggleableMixin} interface.</summary>
        private IToggleableMixin Toggleables (string controlId) => _ribbonFactory.Toggleables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {ISelectableMixin} interface.</summary>
        private ISelectableMixin Selectables (string controlId) => _ribbonFactory.Selectables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {IImageableMixin} interface.</summary>
        private IImageableMixin  Imageables  (string controlId) => _ribbonFactory.Imageables.GetOrDefault(controlId);

        /// <inheritdoc/>
        public string GetDescription(IRibbonControl Control)
            => Controls(Control?.Id)?.Description ?? Unknown(Control);
        /// <inheritdoc/>
        public bool   GetEnabled(IRibbonControl Control)
            => Controls(Control?.Id)?.IsEnabled ?? false;
        /// <inheritdoc/>
        public string GetKeyTip(IRibbonControl Control)
            => Controls(Control?.Id)?.KeyTip ?? "??";
        /// <inheritdoc/>
        public string GetLabel(IRibbonControl Control)
            => Controls(Control?.Id)?.Label ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetScreenTip(IRibbonControl Control)
            => Controls(Control?.Id)?.ScreenTip ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetSuperTip(IRibbonControl Control)
            => Controls(Control?.Id)?.SuperTip ?? Unknown(Control);
        /// <inheritdoc/>
        public bool   GetVisible(IRibbonControl Control)
            => Controls(Control?.Id)?.IsVisible ?? true;

        /// <inheritdoc/>
        public RdControlSize GetSize(IRibbonControl Control)
            => Sizeables(Control?.Id)?.GetSize() ?? RdControlSize.rdLarge;

        /// <inheritdoc/>
        public object GetImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.GetImage().Image ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetShowImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.GetShowImage() ?? true;
        /// <inheritdoc/>
        public bool   GetShowLabel(IRibbonControl Control)
            => Imageables(Control?.Id)?.GetShowLabel() ?? true;

        /// <inheritdoc/>
        public bool   GetPressed(IRibbonControl Control)
            => Toggleables(Control?.Id)?.GetPressed() ?? false;
        /// <inheritdoc/>
        public void   OnActionToggle(IRibbonControl Control, bool Pressed)
            => Toggleables(Control?.Id)?.OnActionToggle(Pressed);

        /// <inheritdoc/>
        public void   OnAction(IRibbonControl Control)
            => Actionables(Control?.Id)?.Clicked();

        /// <inheritdoc/>
        public string GetSelectedItemId(IRibbonControl Control)
            => Selectables(Control?.Id)?.SelectedItemId;
        /// <inheritdoc/>
        public int    GetSelectedItemIndex(IRibbonControl Control)
            => Selectables(Control?.Id)?.SelectedItemIndex ?? 0;
        /// <inheritdoc/>
        public void   OnActionDropDown(IRibbonControl Control, string SelectedId, int SelectedIndex)
            => Selectables(Control?.Id)?.OnActionDropDown(SelectedId, SelectedIndex);
 
        /// <inheritdoc/>
        public int    GetItemCount(IRibbonControl Control)
            => Selectables(Control?.Id)?.ItemCount ?? 0;
        /// <inheritdoc/>
        public string GetItemId(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemId(Index) ?? "";
        /// <inheritdoc/>
        public string GetItemLabel(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemLabel(Index) ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetItemScreenTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemScreenTip(Index) ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetItemSuperTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemSuperTip(Index) ?? Unknown(Control);

        /// <inheritdoc/>
        public object GetItemImage(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemImage(Index) ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetItemShowImage(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowImage(Index) ?? true;
        /// <inheritdoc/>
        public bool   GetItemShowLabel(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowLabel(Index) ?? true;

        /// <summary>Returns a string as the ID of the supplied control suffixed with ' unknown'.</summary>
        private static string Unknown(IRibbonControl Control) 
            => string.Format(CultureInfo.InvariantCulture, $"'{Control?.Id??""}' unknown");
    }
}
