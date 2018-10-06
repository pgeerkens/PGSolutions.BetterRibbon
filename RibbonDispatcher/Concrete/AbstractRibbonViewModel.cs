using System;
using System.Globalization;
using System.Resources;
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
    public abstract class AbstractRibbonViewModel : IRibbonViewModel, IResourceManager {
        /// <summary>Initializes this instance with the supplied {IRibbonUI}.</summary>
        protected void Initialize(IRibbonUI RibbonUI) 
            => _ribbonFactory = new RibbonFactory(RibbonUI);

        /// <summary>Initializes this instance with the supplied {IRibbonUI} and {IResourceManager}.</summary>
        protected void Initialize(IRibbonUI RibbonUI, IResourceManager ResourceManager)
            => _ribbonFactory = new RibbonFactory(RibbonUI, ResourceManager);

        /// <inheritdoc/>
        public object LoadImage(string imageId) => _ribbonFactory.LoadImage(imageId);

        /// <inheritdoc/>
        public IRibbonFactory RibbonFactory => _ribbonFactory; private RibbonFactory _ribbonFactory;

        /// <inheritdoc/>
        public void Invalidate()                            => _ribbonFactory.Invalidate();
        /// <inheritdoc/>
        public void InvalidateControl(string ControlId)     => _ribbonFactory.InvalidateControl(ControlId);
        /// <inheritdoc/>
        public void InvalidateControlMso(string ControlId)  => _ribbonFactory.InvalidateControlMso(ControlId);
        /// <inheritdoc/>
        public void ActivateTab(string ControlId)           => _ribbonFactory.ActivateTab(ControlId);

        #region IRibbonCommon implementation
        /// <summary>All of the defined controls.</summary>
        private IRibbonCommon Controls (string controlId) => _ribbonFactory.Controls.GetOrDefault(controlId);
        /// <inheritdoc/>
        public string GetDescription(IRibbonControl Control)
            => Controls(Control?.Id)?.Description ?? Unknown(Control, "Description");
        /// <inheritdoc/>
        public bool   GetEnabled(IRibbonControl Control)
            => Controls(Control?.Id)?.IsEnabled ?? false;
        /// <inheritdoc/>
        public string GetKeyTip(IRibbonControl Control)
            => Controls(Control?.Id)?.KeyTip ?? "";
        /// <inheritdoc/>
        public string GetLabel(IRibbonControl Control)
            => Controls(Control?.Id)?.Label ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetScreenTip(IRibbonControl Control)
            => Controls(Control?.Id)?.ScreenTip ?? Unknown(Control, "ScreenTip");
        /// <inheritdoc/>
        public string GetSuperTip(IRibbonControl Control)
            => Controls(Control?.Id)?.SuperTip ?? Unknown(Control, "SuperTip");
        /// <inheritdoc/>
        public bool   GetVisible(IRibbonControl Control)
            => Controls(Control?.Id)?.IsVisible ?? true;
        #endregion

        #region ISizeableMixin implementation
        /// <summary>All of the defined controls implementing the {ISizeableMixin} interface.</summary>
        private ISizeableMixin Sizeables(string controlId) => _ribbonFactory.Sizeables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public RdControlSize GetSize(IRibbonControl Control)
            => Sizeables(Control?.Id)?.GetSize() ?? RdControlSize.rdLarge;
        #endregion

        #region IImageableMixin implementation
        /// <summary>All of the defined controls implementing the {IImageableMixin} interface.</summary>
        private IImageableMixin Imageables (string controlId) => _ribbonFactory.Imageables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public object GetImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.GetImage().Image ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetShowImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.GetShowImage() ?? true;
        /// <inheritdoc/>
        public bool   GetShowLabel(IRibbonControl Control)
            => Imageables(Control?.Id)?.GetShowLabel() ?? true;
        #endregion

        #region IToggleableMixin implementation
        /// <summary>All of the defined controls implementing the {IToggleableMixin} interface.</summary>
        private IToggleableMixin Toggleables(string controlId) => _ribbonFactory.Toggleables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public bool   GetPressed(IRibbonControl Control)
            => Toggleables(Control?.Id)?.GetPressed() ?? false;
        /// <inheritdoc/>
        public void   OnActionToggle(IRibbonControl Control, bool Pressed)
            => Toggleables(Control?.Id)?.OnActionToggle(Pressed);
        #endregion

        #region IClickableMixin implementation
        /// <summary>All of the defined controls implementing the {IClickableMixin} interface.</summary>
        private IClickableMixin Actionables(string controlId) => _ribbonFactory.Actionables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public void   OnAction(IRibbonControl Control)
            => Actionables(Control?.Id)?.Clicked();
        #endregion

        #region ISelectableMixin implementation
        /// <summary>All of the defined controls implementing the {ISelectableMixin} interface.</summary>
        private ISelectableMixin Selectables (string controlId) => _ribbonFactory.Selectables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public int    GetItemCount(IRibbonControl Control)
            => Selectables(Control?.Id)?.ItemCount ?? 0;
        /// <inheritdoc/>
        public string GetItemId(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemId(Index) ?? "";
        /// <inheritdoc/>
        public object GetItemImage(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemImage(Index) ?? "MacroSecurity";
        /// <inheritdoc/>
        public string GetItemLabel(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemLabel(Index) ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetItemScreenTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemScreenTip(Index) ?? Unknown(Control);
        /// <inheritdoc/>
        public bool   GetItemShowImage(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowImage(Index) ?? true;
        /// <inheritdoc/>
        public bool   GetItemShowLabel(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowLabel(Index) ?? true;
        /// <inheritdoc/>
        public string GetItemSuperTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemSuperTip(Index) ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetSelectedItemId(IRibbonControl Control)
            => Selectables(Control?.Id)?.SelectedItemId;
        /// <inheritdoc/>
        public int    GetSelectedItemIndex(IRibbonControl Control)
            => Selectables(Control?.Id)?.SelectedItemIndex ?? 0;
        /// <inheritdoc/>
        public void   OnActionDropDown(IRibbonControl Control, string SelectedId, int SelectedIndex)
            => Selectables(Control?.Id)?.OnActionDropDown(SelectedId, SelectedIndex);
        #endregion

        /// <inheritdoc/>
        public IRibbonTextLanguageControl GetControlStrings(string ControlId) =>
            new RibbonTextLanguageControl(
                    GetCurrentUIString($"{ControlId}_Label")            ?? Unknown(ControlId),
                    GetCurrentUIString($"{ControlId}_ScreenTip")        ?? Unknown(ControlId, "ScreenTip"),
                    GetCurrentUIString($"{ControlId}_SuperTip")         ?? Unknown(ControlId, "SuperTip"),
                    GetCurrentUIString($"{ControlId}_KeyTip")           ?? "",
                    GetCurrentUIString($"{ControlId}_AlternativeLabel") ?? Unknown(ControlId, "Alternate"),
                    GetCurrentUIString($"{ControlId}_Description")      ?? Unknown(ControlId, "Description")
            );

        /// <inheritdoc/>
        public object GetImage(string Name) => ResourceManager.Value.GetResourceImage(Name);

        /// <summary>Returns a string as the ID of the supplied control suffixed with ' Unknown'.</summary>
        protected static string Unknown(IRibbonControl Control) => Unknown(Control?.Id, "Unknown");

        /// <summary>Returns a string as the ID of the supplied control suffixed with ' Unknown'.</summary>
        protected static string Unknown(string controlId)       => Unknown(controlId, "Unknown");

        /// <summary>Returns a string as the ID of the supplied control suffixed with the supplied string.</summary>
        protected static string Unknown(IRibbonControl Control, string suffix) => Unknown(Control?.Id, suffix);

        /// <summary>Returns a string as the ID of the supplied control suffixed with the supplied string.</summary>
        protected static string Unknown(string controlId, string suffix)
            => string.Format(CultureInfo.InvariantCulture, $"'{controlId ?? ""}' {suffix}");

        /// <summary>TODO</summary>
        protected abstract Lazy<ResourceManager> ResourceManager {  get; }

        private string GetCurrentUIString(string controlId) => ResourceManager.Value.GetCurrentUIString(controlId);
    }
}
