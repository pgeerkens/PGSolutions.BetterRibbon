////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
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
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226"> For more information about adding callback methods.</a>
    /// 
    /// Whenever possible the AbstractDispatcher will return default values acceptable to OFFICE
    /// even if the Control.ControlId supplied to a callback is unknown. These defaults are
    /// chosen to maximize visibility for the unknown control, but disable its functionality.
    /// This is believed to support the principle of 'least surprise', given the OFFICE 
    /// Ribbon's propensity to fail, silently and/or fatally, at the slightest provocation.
    /// 
    /// This class must be COM-Visible for the Ribbon callbacks to be received!
    /// </remarks>
    [Description("Implementation of (all) the callbacks for the Fluent Ribbon.")]
    [Serializable, ComVisible(true), CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICallbackDispatcher))]
    [Guid(Guids.AbstractDispatcher)]
    public abstract class AbstractDispatcher:  ICallbackDispatcher {
        protected AbstractDispatcher(){ }

        /// <inheritdoc/>
        public ViewModelFactory ViewModelFactory { get; private set; }

        protected void SetViewModelFactory(ViewModelFactory factory) {
            ViewModelFactory?.ClearChangedListeners();
            ViewModelFactory = factory;
            ViewModelFactory.Changed += OnPropertyChanged;
            RibbonUI?.Invalidate();
        }

        public virtual void RegisterWorkbook(string workbookName) { }

        protected virtual void OnPropertyChanged(object sender, IControlChangedEventArgs e)
        => RibbonUI?.InvalidateControl(e?.Control.ControlId);

        /// summary>.<summary/>
        public virtual object LoadImage(string ImageId) => ResourceLoader.GetImage(ImageId);

        #region IRibbonExtensibility implementation
        /// <summary>Raised to signal completion of the Ribbon load.</summary>
        public event EventHandler Initialized;

        /// summary>.<summary/>
        public             IRibbonUI       RibbonUI       { get; private set; }

        /// summary>.<summary/>
        protected abstract string          RibbonXml      { get; }

        /// <summary>The <see cref="IResourceLoader"/> for common shared resources.</summary>
        public    abstract IResourceLoader ResourceLoader { get; }

        /// <summary>The callback from VSTO/VSTA requesting the Ribbon XML text.</summary>
        /// <param name="RibbonID"></param>
        /// <returns>Returns the supplied RibbonXml after parsing it to creates the <see cref="RibbonViewModel"/>.</returns>
        [SuppressMessage("Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "RibbonID")]
        public virtual string GetCustomUI(string RibbonID)  {
            SetViewModelFactory(RibbonXml.ParseXmlTabs());

            return RibbonXml;
        }

        /// <summary>Callback from VSTO/VSTA signalling successful Ribbon load, and providing the <see cref="IRibbonUI"/> handle.</summary>
        public virtual void OnRibbonLoad(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;

            Initialized?.Invoke(this, EventArgs.Empty);
        }
        #endregion

        #region IControlVM implementation
        /// <summary>All of the defined controls.</summary>
        private IControlVM Controls (string controlId) => ViewModelFactory.Controls.GetOrDefault(controlId);
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
            => Imageables(Control?.Id)?.Image?.Image() ?? "MacroSecurity";
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
        private ISelectItemsVM SelectItems (string controlId) => ViewModelFactory.SelectItems.GetOrDefault(controlId);
        /// <inheritdoc/>
        public int    GetItemCount(IRibbonControl Control)
        => SelectItems(Control?.Id)?.Items?.Count ?? 0;
        /// <inheritdoc/>
        public string GetItemId(IRibbonControl Control, int Index)
        => SelectItems(Control?.Id)?.Items[Index].ControlId ?? "";
        /// <inheritdoc/>
        public object GetItemImage(IRibbonControl Control, int Index) {
            var image = SelectItems(Control?.Id)?.Items[Index].Image;
            return image.IsMso ? image.ImageMso : (object)image.ImageDisp ?? "MacroSecurity";
        }
        /// <inheritdoc/>
        public string GetItemLabel(IRibbonControl Control, int Index)
        => SelectItems(Control?.Id)?.Items[Index].Label ?? Control.Unknown();
        /// <inheritdoc/>
        public string GetItemScreenTip(IRibbonControl Control, int Index)
        => SelectItems(Control?.Id)?.Items[Index].ScreenTip ?? Control.Unknown();
        /// <inheritdoc/>
        public string GetItemSuperTip(IRibbonControl Control, int Index)
        => SelectItems(Control?.Id)?.Items[Index].SuperTip ?? Control.Unknown();
        #endregion

        #region ISelectable2VM implementation - DropDown
        /// <summary>All of the defined controls implementing the {ISelectableMixin} interface.</summary>
        private ISelectablesVM Selectables2(string controlId) => ViewModelFactory.Selectable2.GetOrDefault(controlId);

        /// <inheritdoc/>
        public string GetSelectedItemId(IRibbonControl control)
        => Selectables2(control?.Id)?.SelectedItemId;
        /// <inheritdoc/>
        public int    GetSelectedItemIndex(IRibbonControl control)
            => Selectables2(control?.Id)?.SelectedItemIndex ?? 0;

        /// <inheritdoc/>
        public void   OnActionSelected(IRibbonControl control, string selectedId, int selectedIndex)
        => Selectables2(control?.Id)?.OnSelectionMade(control, selectedId, selectedIndex);
        #endregion

        #region GallerySizeVM implementation
        /// <summary>All of the defined controls implementing the {IClickableVM} interface.</summary>
        private IGallerySizeVM GallerySizes(string controlId) => ViewModelFactory.GallerySizes.GetOrDefault(controlId);

        public int GetItemHeight(IRibbonControl control) => GallerySizes(control?.Id)?.ItemHeight ?? 0;

        public int GetItemWidth(IRibbonControl control) => GallerySizes(control?.Id)?.ItemWidth ?? 0;
        #endregion

        #region IEditableVM implementation - EditBox & ComboBox
        /// <summary>All of the defined controls implementing the {IClickableVM} interface.</summary>
        private IEditableVM TextEditables(string controlId) => ViewModelFactory.Editables.GetOrDefault(controlId);

        public string GetText(IRibbonControl control)
        => TextEditables(control?.Id)?.Text ?? "";

        public void   OnTextChanged(IRibbonControl control, string text)
        => TextEditables(control?.Id)?.OnEdited(control, text);
        #endregion

        #region IDynamicMenuVM implementation
        /// <summary>All of the defined controls implementing the {IClickableVM} interface.</summary>
        private IDynamicMenuVM DynamicMenus(string controlId) => ViewModelFactory.DynamicMenus.GetOrDefault(controlId);

        public string GetContent(IRibbonControl control) {
            var content = @"<menu xmlns=\'http://schemas.microsoft.com/office/2006/01/customui\'></menu>";
            DynamicMenus(control?.Id)?.OnGetContent(control, out content);
            return content;
        }
        #endregion

        #region IDescriptionableVM implementation
        /// <summary>All of the defined controls.</summary>
        private IDescriptionableVM Descriptionables(string controlId) => ViewModelFactory.Descriptionables.GetOrDefault(controlId);

        /// <inheritdoc/>
        public string GetDescription(IRibbonControl Control)
            => Descriptionables(Control?.Id)?.Description ?? Control.Unknown("Description");
        #endregion

        #region IMenuSeparatorVM implementation
        private IMenuSeparatorVM MenuSeparators(string controlId) => ViewModelFactory.MenuSeparators.GetOrDefault(controlId);

        public string GetTitle(IRibbonControl control) =>MenuSeparators(control?.Id)?.Title ?? "";
        #endregion
    }
}
