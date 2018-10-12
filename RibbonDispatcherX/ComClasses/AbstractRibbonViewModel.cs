////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ControlMixins;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.RibbonDispatcher.ComClasses {

    /// <summary>Implementation of (all) the callbacks for the Fluent Ribbon; for .NET clients.</summary>
    /// <remarks>
    /// DOT NET clients are expected to find it more convenient to inherit their ViewModel 
    /// class from {AbstractRibbonViewModel} than to compose against an instance of 
    /// {RibbonViewModel}. COM clients will most likely find the reverse true. 
    /// 
    /// The callback names are chosen to be identical to the corresponding xml tag in the
    /// Ribbon schema, except for:
    ///  - PascalCase instead of camelCase; and
    ///  - In some instances, a disambiguating usage-suffix such as in OnActionToggle(,)
    ///    instead of a plain OnAction(,).
    ///    
    /// Whenever possible the ViewModel will return default values acceptable to OFFICE
    /// even if the Control.Id supplied to a callback is unknown. These defaults are
    /// chosen to maximize visibility for the unknown control, but disable its functionality.
    /// This is believed to support the principle of 'least surprise', given the OFFICE 
    /// Ribbon's propensity to fail, silently and/or fatally, at the slightest provocation.
    /// </remarks>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [ComDefaultInterface(typeof(IRibbonViewModel))]
    [Guid(Guids.AbstractDispatcher)]
    public abstract class AbstractRibbonViewModel : IRibbonViewModel {
        /// <summary>Initializes this instance with the supplied {IRibbonUI}.</summary>
        protected AbstractRibbonViewModel() : this(new RibbonFactory()) { }

        /// <summary>Initializes this instance with the supplied {IRibbonUI} and {IResourceManager}.</summary>
        protected AbstractRibbonViewModel(IResourceManager ResourceManager) : this(new RibbonFactory(ResourceManager)) { }

        private AbstractRibbonViewModel(RibbonFactory ribbonFactory) {
            _ribbonFactory = ribbonFactory;
            _ribbonFactory.Changed += PropertyChanged;
        }

        public virtual void OnRibbonLoad(IRibbonUI ribbonUI) => RibbonUI = ribbonUI;

        /// <inheritdoc/>
        public object LoadImage(string imageId) => _ribbonFactory.LoadImage(imageId);

        /// <inheritdoc/>
        public IRibbonFactory RibbonFactory => _ribbonFactory; private RibbonFactory _ribbonFactory;

        public IRibbonUI RibbonUI { get; protected set; }

        private void PropertyChanged(object sender, IControlChangedEventArgs e) => RibbonUI?.InvalidateControl(e.ControlId);

        /// <inheritdoc/>
        public void Invalidate()                                => RibbonUI?.Invalidate();
        /// <inheritdoc/>
        public void InvalidateControl(string ControlId)         => RibbonUI?.InvalidateControl(ControlId);
        /// <inheritdoc/>
        public void InvalidateControlMso(string ControlId)      => RibbonUI?.InvalidateControlMso(ControlId);
        /// <inheritdoc/>
        public void ActivateTab(string ControlId)               => RibbonUI?.ActivateTab(ControlId);
        /// <inheritdoc/>
        public void ActivateTabQ(string ControlId, string ns)   => RibbonUI?.ActivateTabQ(ControlId, ns);

        #region IRibbonCommon implementation
        /// <summary>All of the defined controls.</summary>
        private IRibbonCommon Controls (string controlId) => _ribbonFactory.Controls.GetOrDefault(controlId);
        /// <inheritdoc/>
        public string GetDescription(IRibbonControl Control)
            => Controls(Control?.Id)?.Description ?? Control.Unknown("Description");
        /// <inheritdoc/>
        public bool   GetEnabled(IRibbonControl Control)
            => Controls(Control?.Id)?.IsEnabled ?? false;
        /// <inheritdoc/>
        public string GetKeyTip(IRibbonControl Control)
            => Controls(Control?.Id)?.KeyTip ?? "";
        /// <inheritdoc/>
        public string GetLabel(IRibbonControl Control)
            => Controls(Control?.Id)?.Label ?? Control.Unknown();
        /// <inheritdoc/>
        public string GetScreenTip(IRibbonControl Control)
            => Controls(Control?.Id)?.ScreenTip ?? Control.Unknown("ScreenTip");
        /// <inheritdoc/>
        public string GetSuperTip(IRibbonControl Control)
            => Controls(Control?.Id)?.SuperTip ?? Control.Unknown("SuperTip");
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
        private IClickableMixin Actionables(string controlId) => _ribbonFactory.Clickables.GetOrDefault(controlId);
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
            => Selectables(Control?.Id)?.ItemLabel(Index) ?? Control.Unknown();
        /// <inheritdoc/>
        public string GetItemScreenTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemScreenTip(Index) ?? Control.Unknown();
        /// <inheritdoc/>
        public bool   GetItemShowImage(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowImage(Index) ?? true;
        /// <inheritdoc/>
        public bool   GetItemShowLabel(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowLabel(Index) ?? true;
        /// <inheritdoc/>
        public string GetItemSuperTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemSuperTip(Index) ?? Control.Unknown();
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
    }
}
