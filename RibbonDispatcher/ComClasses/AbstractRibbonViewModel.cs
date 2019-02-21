////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;

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

        /// <summary>Initializes this instance with the supplied {IRibbonUI} and {IResourceManager}.</summary>
        protected AbstractRibbonViewModel(string controlId, IResourceManager ResourceManager)
        : this(controlId, new RibbonFactory(ResourceManager)) { }

        private AbstractRibbonViewModel(string controlId, RibbonFactory ribbonFactory) {
            Id             = controlId;
            _ribbonFactory = ribbonFactory;
            _ribbonFactory.Changed += PropertyChanged;
        }

        #region IRibbonExtensibility implementation
        /// <summary>Raised to signal completion of the Ribbon load.</summary>
        public event EventHandler Initialized;

        /// <summary>The callback from VSTO/VSTA requesting the Ribbon XML text.</summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "RibbonID")]
        public string GetCustomUI(string RibbonID) => ParseXml(RibbonXml);

        /// <summary>Callback from VSTO/VSTA signalling successful Ribbon load, and providing the <see cref="IRibbonUI"/> handle.</summary>
        [CLSCompliant(false)]
        public virtual void OnRibbonLoad(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;

            Initialized?.Invoke(this, EventArgs.Empty);

            Invalidate();
        }

        protected abstract string RibbonXml { get; }

        /// <summary>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</summary>
        /// <param name="ribbonXml"></param>
        private string ParseXml(string ribbonXml) {
            XNamespace mso2009 = "http://schemas.microsoft.com/office/2009/07/customui";
            foreach (var group in XDocument.Parse(ribbonXml).Root.Descendants(mso2009+"group")) {
                var viewModel = AddGroupViewModel(group.Attribute("id").Value);

                foreach (var element in group.Descendants()) {
                    switch (element.Name) {
                        case XName name when name == mso2009+"toggleButton":
                            viewModel.Add<IRibbonToggleSource>(RibbonFactory.NewRibbonToggle(element.Attribute("id").Value));
                            break;
                        case XName name when name == mso2009+"checkBox":
                            viewModel.Add<IRibbonToggleSource>(RibbonFactory.NewRibbonCheckBox(element.Attribute("id").Value));
                            break;
                        case XName name when name == mso2009+"dropDown":
                            viewModel.Add<IRibbonDropDownSource>(RibbonFactory.NewRibbonDropDown(element.Attribute("id").Value));
                            break;
                        case XName name when name == mso2009+"button":
                            viewModel.Add<IRibbonButtonSource>(RibbonFactory.NewRibbonButton(element.Attribute("id").Value));
                            break;
                        default:
                            break;
                    }
                }
            }

            return ribbonXml;
        }
        #endregion

        /// <summary>Registers and returns a new <see cref="RibbonGroupViewModel"/> as named.</summary>
        public virtual RibbonGroupViewModel AddGroupViewModel(string groupName) {
            var viewModel = RibbonFactory.NewRibbonGroup(groupName);
            _groupViewModels.Add(viewModel);
            return viewModel;
        }

        public IReadOnlyList<RibbonGroupViewModel> GroupViewModels => _groupViewModels.AsReadOnly();

        private List<RibbonGroupViewModel> _groupViewModels { get; }
                                        = new List<RibbonGroupViewModel>();

        /// <inheritdoc/>
        public object LoadImage(string imageId) => _ribbonFactory.LoadImage(imageId);

        /// <inheritdoc/>
        public IRibbonFactory RibbonFactory => _ribbonFactory; private RibbonFactory _ribbonFactory;

        public IRibbonUI RibbonUI { get; protected set; }

        private void PropertyChanged(object sender, IControlChangedEventArgs e) => RibbonUI?.InvalidateControl(e.ControlId);

        /// <inheritdoc/>
        public void Invalidate()                                => RibbonUI?.Invalidate();
        /// <inheritdoc/>
        public void InvalidateTab()                             => RibbonUI?.InvalidateControl(Id);
        /// <inheritdoc/>
        public void InvalidateControl(string ControlId)         => RibbonUI?.InvalidateControl(ControlId);
        /// <inheritdoc/>
        public void InvalidateControlMso(string ControlId)      => RibbonUI?.InvalidateControlMso(ControlId);
        /// <inheritdoc/>
        public void ActivateTab(string ControlId)               => RibbonUI?.ActivateTab(ControlId);
        /// <inheritdoc/>
        public void ActivateTabQ(string ControlId, string ns)   => RibbonUI?.ActivateTabQ(ControlId, ns);

        protected string Id { get; }

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

        #region ISizeable implementation
        /// <summary>All of the defined controls implementing the {ISizeable} interface.</summary>
        private ISizeable Sizeables(string controlId) => _ribbonFactory.Sizeables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public bool GetSize(IRibbonControl Control)
            => (Sizeables(Control?.Id)?.IsLarge ?? true);
        #endregion

        #region IImageable implementation
        /// <summary>All of the defined controls implementing the {IImageable} interface.</summary>
        private IImageable Imageables (string controlId) => _ribbonFactory.Imageables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public object GetImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.Image.Image ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetShowImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.ShowImage ?? false;
        /// <inheritdoc/>
        public bool   GetShowLabel(IRibbonControl Control)
            => Imageables(Control?.Id)?.ShowLabel ?? true;
        #endregion

        #region IToggleable implementation
        /// <summary>All of the defined controls implementing the {IToggleable} interface.</summary>
        private IToggleable Toggleables(string controlId) => _ribbonFactory.Toggleables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public bool   GetPressed(IRibbonControl Control)
            => Toggleables(Control?.Id)?.IsPressed ?? false;
        /// <inheritdoc/>
        public void   OnActionToggle(IRibbonControl Control, bool Pressed)
            => Toggleables(Control?.Id)?.OnToggled(this, Pressed);
        #endregion

        #region IClickable implementation
        /// <summary>All of the defined controls implementing the {IClickable} interface.</summary>
        private IClickable Actionables(string controlId) => _ribbonFactory.Clickables.GetOrDefault(controlId);
 
        /// <inheritdoc/>
        public void   OnAction(IRibbonControl Control)   => Actionables(Control?.Id)?.OnClicked(this, EventArgs.Empty);
        #endregion

        #region ISelectableMixin implementation
        /// <summary>All of the defined controls implementing the {ISelectableMixin} interface.</summary>
        private ISelectable Selectables (string controlId) => _ribbonFactory.Selectables.GetOrDefault(controlId);
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
