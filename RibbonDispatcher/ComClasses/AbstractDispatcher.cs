////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IGroupList = IReadOnlyList<GroupVM>;

    /// <summary>Implementation of (all) the callbacks for the Fluent Ribbon; for .NET clients.</summary>
    /// <remarks>
    /// DOT NET clients are expected to find it more convenient to inherit their ViewModel 
    /// class from {AbstractDispatcher} than to compose against an instance of 
    /// {RibbonViewModel}. COM clients will most likely find the reverse true. 
    /// 
    /// The callback names are chosen to be identical to the corresponding xml tag in the
    /// Ribbon schema, except for:
    ///  - PascalCase instead of camelCase; and
    ///  - In some instances, a disambiguating usage-suffix such as in OnActionToggle(,)
    ///    instead of a plain OnAction(,).
    ///    
    /// Whenever possible the ViewModel will return default values acceptable to OFFICE
    /// even if the Control.ControlId supplied to a callback is unknown. These defaults are
    /// chosen to maximize visibility for the unknown control, but disable its functionality.
    /// This is believed to support the principle of 'least surprise', given the OFFICE 
    /// Ribbon's propensity to fail, silently and/or fatally, at the slightest provocation.
    /// 
    /// This class must be COM-Visible for the Ribbon callbacks to be received!
    /// </remarks>
    [Description("Implementation of (all) the callbacks for the Fluent Ribbon; for .NET clients.")]
    [Serializable, ComVisible(true), CLSCompliant(true)]
    [ComDefaultInterface(typeof(ICallbackDispatcher))]
    [Guid(Guids.AbstractDispatcher)]
    public abstract class AbstractDispatcher: ICallbackDispatcher, IRibbonViewModel {

        /// <summary>Initializes this instance with the supplied {IRibbonUI} and {IResourceManager}.</summary>
        protected AbstractDispatcher(string controlId, IResourceManager resourceManager){
            ControlId        = controlId;
            ViewModelFactory = new ViewModelFactory(resourceManager);
            ViewModelFactory.Changed += OnPropertyChanged;
        }

        /// <inheritdoc/>
        public   string           ControlId        { get; }

        /// <inheritdoc/>
        public   ViewModelFactory ViewModelFactory { get; }

        /// <inheritdoc/>
        public   IRibbonUI        RibbonUI         { get; private set; }

        public IGroupVM GetGroup(string groupId)
        => GroupViewModels.FirstOrDefault(vm => vm.Id == groupId);

        private void OnPropertyChanged(object sender, IControlChangedEventArgs e)
        => RibbonUI?.InvalidateControl(e.ControlId);

        #region IRibbonExtensibility implementation
        /// <inheritdoc/>
        private IGroupList GroupViewModels { get; set; }

        /// <summary>Raised to signal completion of the Ribbon load.</summary>
        public event EventHandler Initialized;

        /// <summary>The callback from VSTO/VSTA requesting the Ribbon XML text.</summary>
        /// <param name="RibbonID"></param>
        /// <returns>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</returns>
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "RibbonID")]
        public string GetCustomUI(string RibbonID) {
            GroupViewModels = ViewModelFactory.ParseXml(RibbonXml);

            return RibbonXml;
        }

        /// <summary>Callback from VSTO/VSTA signalling successful Ribbon load, and providing the <see cref="IRibbonUI"/> handle.</summary>
        [CLSCompliant(false)]
        public virtual void OnRibbonLoad(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;

            Initialized?.Invoke(this, EventArgs.Empty);

            this.InvalidateTab();
        }

        protected abstract string RibbonXml { get; }

        /// <inheritdoc/>
        public object LoadImage(string ImageId) => ViewModelFactory.LoadImage(ImageId);
        #endregion

        #region IControlVM implementation
        /// <summary>All of the defined controls.</summary>
        private IControlVM Controls (string controlId) => ViewModelFactory.Controls.GetOrDefault(controlId);
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

        #region ISizeableVM implementation
        /// <summary>All of the defined controls implementing the {ISizeableVM} interface.</summary>
        private ISizeableVM Sizeables(string controlId) => ViewModelFactory.Sizeables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public bool GetSize(IRibbonControl Control)
            => (Sizeables(Control?.Id)?.IsLarge ?? true);
        #endregion

        #region IImageableVM implementation
        /// <summary>All of the defined controls implementing the {IImageableVM} interface.</summary>
        private IImageableVM Imageables (string controlId) => ViewModelFactory.Imageables.GetOrDefault(controlId);
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

        #region IClickableVM implementation
        /// <summary>All of the defined controls implementing the {IClickableVM} interface.</summary>
        private IClickableVM Clickables(string controlId) => ViewModelFactory.Clickables.GetOrDefault(controlId);

        /// <inheritdoc/>
        public void OnAction(IRibbonControl Control) => Clickables(Control?.Id)?.OnClicked(Control);
        #endregion

        #region IToggleableVM implementation
        /// <summary>All of the defined controls implementing the {IToggleableVM} interface.</summary>
        private IToggleableVM Toggleables(string controlId) => ViewModelFactory.Toggleables.GetOrDefault(controlId);
        /// <inheritdoc/>
        public bool   GetPressed(IRibbonControl Control)
            => Toggleables(Control?.Id)?.IsPressed ?? false;
        /// <inheritdoc/>
        public void   OnActionToggle(IRibbonControl Control, bool IsPressed)
            => Toggleables(Control?.Id)?.OnToggled(Control, IsPressed);
        #endregion

        #region ISelectableVM implementation - DropDown & ComboBox
        /// <summary>All of the defined controls implementing the {ISelectableMixin} interface.</summary>
        private ISelectableVM Selectables (string controlId) => ViewModelFactory.Selectables.GetOrDefault(controlId);
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
        public string GetItemSuperTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemSuperTip(Index) ?? Control.Unknown();
        #endregion

        #region ISelectable2VM implementation - DropDown
        /// <summary>All of the defined controls implementing the {ISelectableMixin} interface.</summary>
        private ISelectable2VM Selectables2(string controlId) => ViewModelFactory.Selectables2.GetOrDefault(controlId);

        /// <inheritdoc/>
        public string GetSelectedItemId(IRibbonControl Control)
            => Selectables2(Control?.Id)?.SelectedItemId;
        /// <inheritdoc/>
        public int GetSelectedItemIndex(IRibbonControl Control)
            => Selectables2(Control?.Id)?.SelectedItemIndex ?? 0;

        ///// <inheritdoc/>
        //public bool   GetItemShowImage(IRibbonControl Control, int Index)
        //    => Selectables(Control?.Id)?.ItemShowImage(Index) ?? true;
        ///// <inheritdoc/>
        //public bool   GetItemShowLabel(IRibbonControl Control, int Index)
        //    => Selectables(Control?.Id)?.ItemShowLabel(Index) ?? true;

        /// <inheritdoc/>
        public void   OnActionDropDown(IRibbonControl Control, string SelectedId, int SelectedIndex)
            => Selectables2(Control?.Id)?.OnSelectionMade(Control, SelectedId, SelectedIndex);
        #endregion

        #region IEditableVM implementation - EditBox & ComboBox
        /// <summary>All of the defined controls implementing the {IClickableVM} interface.</summary>
        private IEditableVM TextEditables(string controlId) => ViewModelFactory.TextEditables.GetOrDefault(controlId);

        public string GetText(IRibbonControl control)
        => TextEditables(control?.Id)?.Text ?? "";

        public void   OnTextChanged(IRibbonControl control, string text)
        => TextEditables(control?.Id)?.OnEdited(control, text);
        #endregion
    }
}
